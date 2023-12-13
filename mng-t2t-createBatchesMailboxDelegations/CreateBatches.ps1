<#
  Date: 22/05/2023
  Environment: Unishare migration NL - Medlon Env.
  Author: Diego Ordonez
  Synopsys: Create batches on the basis bellow:
    - The purpose of this script is group the users and their delegations in the same batch
    - Avoid cohexistences whenever it's possible (User migrated and mailbox not migrated)
    - Identify relathinships between main owner and their delegates
    - Provide the capacity to set the size for each wave
#>


$PATH_ROOT = $PSScriptRoot
$WAVE_SIZE = 200
$TOLERANCE_SIZE = 10
$PATH_PERMISSIONS = "$($PATH_ROOT)\AllPermissions.csv"
$PATH_MAILBOXINFO = "$($PATH_ROOT)\AllMailboxesInfo.csv"
$PATH_AADINFO = "$($PATH_ROOT)\UserDepartment.csv"
$PATH_EXPORT = "$($PATH_ROOT)\export"
$TODAY = Get-Date -Format "yyyyMMddHHmm"
$PATH_OUT = "$($PATH_EXPORT)\$($TODAY)_NewMigrationBatches.csv"


function New-BatchObject {
  param(
    [Parameter(Mandatory = $true)]$UserPrincipalName,
    [Parameter(Mandatory = $false)]$Delegations = @(),
    [Parameter(Mandatory = $false)]$LinkCount = 0
  )
  $ThisBatchObject = [PSCustomObject]@{
    UserPrincipalName = $UserPrincipalName
    LinkCount         = $LinkCount
    MainLink          = ""
    Delegations       = $Delegations
    BatchId           = ""
    Department        = ($AADInfo | Where-Object { $_.UserPrincipalName -eq $UserPrincipalName }).Department
    Allocated         = $false

  }
  return $ThisBatchObject
}

function Measure-links {
  param(
    [Parameter(Mandatory = $true)]$ListPermissions
  )

  return $ListPermissions | Group-Object -Property MailboxOwnerUPN | Sort-Object Count -Descending
}

function Get-UniqueValueFromLinks {
  param(
    [Parameter(Mandatory = $true)]$BatchObject
  )

  $BatchObject.Delegations = $BatchObject.Delegations | Select-Object -Unique User

  #Set new linkCount
  $BatchObject.linkCount = ($BatchObject.Delegations | Measure-Object).Count

  return $BatchObject
}

#This function will remove non migrable delegations
function Optimize-Links {
  param(
    [Parameter(Mandatory = $true)]$BatchObject
  )

  #Count removable links
  $totalLinksToRemove = ($BatchObject.Delegations | Where-Object { ($_.User -like "ExchangePublishedUser*") -or ($_.User -like "NT:S*") } | Measure-Object).Count


  #Remove external delegations
  $BatchObject.Delegations = $BatchObject.Delegations | Where-Object { $_.User -notlike "ExchangePublishedUser*" }
  #Remove orphan delegations
  $BatchObject.Delegations = $BatchObject.delegations | Where-Object { $_.User -notlike "NT:S*" }
  
  #remove references from linkCount
  $BatchObject.LinkCount = $BatchObject.LinkCount - $totalLinksToRemove

  return $BatchObject

}

function Get-DelegateUPN {
  param(
    [Parameter(Mandatory = $true)]$BatchObject
  )

  $uPNsFound = @()

  $BatchObject.Delegations | Where-Object { $_.User -notlike "Predeterminado" } | ForEach-Object {

    $thisDN = $_.User

    $thisMailboxInfo = $MailboxInfo | Where-Object { $_.DisplayName -like $thisDN }

    $uPNsFound += $thisMailboxInfo

  }

  $BatchObject.Delegations = $uPNsFound

  $BatchObject.linkCount = ($uPNsFound | Measure-Object).Count

  return $BatchObject
}

function New-BatchAssignment {
  param(
    [Parameter(Mandatory = $true)]$UsersPackage,
    [Parameter(Mandatory = $true)]$GlobalStore,
    [Parameter(Mandatory = $true)]$BatchId
  )

  Write-Host "New batch assignment"

  $nextHop = [PSCustomObject]@{
    BatchId  = $BatchId
    Continue = $false
    Store    = @()
  }


  #add UserPackageHeader
  $headerBatchObject = New-BatchObject -UserPrincipalName $UsersPackage.UserPrincipalName
  $headerBatchObject.BatchID = $BatchId
  $headerBatchObject.Allocated = $true
  $headerBatchObject.MainLink = "Self"

  #Check if its the first iteration

  if ((($GlobalStore | Measure-Object).Count -eq 0) -and $BatchId -eq 1) {
    $GlobalStore += $headerBatchObject
  }
  else {

    if (($GlobalStore.UserPrincipalName.contains($thisUPN)) -eq $false) {
  
      $GlobalStore += $headerBatchObject
    }
  }



  #IInclud in the same batch as header. It creates the link between header and delegates
  $UsersPackage.Delegations | ForEach-Object {

    $thisUPN = $_.UserPrincipalName

    if (($GlobalStore.UserPrincipalName.contains($thisUPN)) -eq $false) {
      
      $objectBatch = New-BatchObject -UserPrincipalName $thisUPN
      $objectBatch.MainLink = $headerBatchObject.UserPrincipalName
      $objectBatch.BatchID = $BatchId
      $objectBatch.Allocated = $true
      $GlobalStore += $objectBatch

    }

  }

  $nextHop.Store = $GlobalStore

  $currentBatchMeasure = ($GlobalStore | Group-Object BatchId | Where-Object { $_.Name -eq $BatchId }).Count

  if ($currentBatchMeasure -le $WAVE_SIZE) {
    $nextHop.Continue = $true
  }
  
  return $nextHop

}



#main

$PermissionsAll = Import-Csv -Path $PATH_PERMISSIONS -Encoding utf8
$MailboxInfo = Import-Csv -Path $PATH_MAILBOXINFO -Encoding utf8
$AADInfo = Import-Csv -Path $PATH_AADINFO -Encoding utf8

#1. Group Permissions and measure links
$PermissionsAll = Measure-links -ListPermissions $PermissionsAll


#2. Create Objects
$allUsersToBatch = @()

$PermissionsAll | ForEach-Object {

  #Isolate delegations
  $thisDelegations = $_.Group

  $allUsersToBatch += New-BatchObject -UserPrincipalName $_.Name -Delegations $thisDelegations -LinkCount $_.Count
}

#3. Optimize delegations

# Unique values
$allUsersToBatch = $allUsersToBatch | ForEach-Object { Get-UniqueValueFromLinks -BatchObject $_ }
#Remove orphans and non this organization users
$allUsersToBatch = $allUsersToBatch | ForEach-Object { Optimize-Links -BatchObject $_ }
# Non longer exists and get UPN
$allUsersToBatch = $allUsersToBatch | ForEach-Object { Get-DelegateUPN -BatchObject $_ }

#3.1 Re-sort results
$allUsersToBatch = $allUsersToBatch | Sort-Object LinkCount -Descending

#4. Split Users in two groups. Mailboxes with dependencies, mailbos with no dependencies

$activeDependencies = $allUsersToBatch | Where-Object { $_.LinkCount -ge 1 }
$nonDependencies = $allUsersToBatch | Where-Object { $_.LinkCount -le 0 }

#5. Start batch creation
$BatchCount = 1
$GlobalBatches = @()

$activeDependencies | ForEach-Object {
  $ThisRelationShip = $_
  $thisBatchAssignment = New-BatchAssignment -GlobalStore $GlobalBatches -UsersPackage $ThisRelationShip -BatchId $BatchCount
  $GlobalBatches = $thisBatchAssignment.Store

  if ($thisBatchAssignment.Continue -eq $false) {
    $BatchCount++
  }

}

$nonDependencies | ForEach-Object {
  $ThisRelationShip = $_
  $thisBatchAssignment = New-BatchAssignment -GlobalStore $GlobalBatches -UsersPackage $ThisRelationShip -BatchId $BatchCount
  $GlobalBatches = $thisBatchAssignment.Store

  if ($thisBatchAssignment.Continue -eq $false) {
    $BatchCount++
  }

}

$GlobalBatches | Group-Object BatchID

$GlobalBatches | Select-Object UserPrincipalName, Department, BatchID, MainLink | Export-Csv -Path $PATH_OUT -NoTypeInformation -Encoding utf8