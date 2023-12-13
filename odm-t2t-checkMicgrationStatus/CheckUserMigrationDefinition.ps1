<#
    Date: 17/01/2023
    Author: Diego Ordonez
    Synopsis: Check user definition schema json again the collected data in common file
#>


$PATH_SCRIPT = $PSScriptRoot
$PATH_CONFIG = "$($PATH_SCRIPT)\project_config.json"
$PATH_USER_DEFINITION = "$($PATH_SCRIPT)\user_profiles.json"
$PATH_TEMPLATES = "$($PATH_SCRIPT)\task_templates.json"
$PATH_COMMON = "$($PSScriptRoot)\common"
$COMMON_FILE = "$($PATH_COMMON)\MigrationGlobal.csv"
$PATH_EXPORT = "$($PSScriptRoot)\exports"
$__CONFIG = Get-Content -Path $PATH_CONFIG | ConvertFrom-Json
$__USERDEF = Get-Content -Path $PATH_USER_DEFINITION | ConvertFrom-Json
$__TASKTEMPLATES = Get-Content -Path $PATH_TEMPLATES | ConvertFrom-Json
$TODAY =  Get-Date


#UserTypes
$USER_PROFILES = $__USERDEF.params.profiles

function Get-DUARequirement{
  param(
    [Parameter(Mandatory = $true)]$MigrationObject
  )

  #Get profileType
  $userProf = $USER_PROFILES | Where-Object {$_.licenseGroup -eq $MigrationObject.licenseGroup}
  
  #1. has expected groupd
  if($null -ne $userProf){
    #2. Store DUA definition from user_prfiles.json
    $MigrationObject.RequiresDUADeployment = $userProf.requiresDUADeployment

    if($MigrationObject.RequiresDUADeployment -eq $false){
      $MigrationObject.ODMAppSwitchTask = "NotMigrable"
    }

  }
  

  return $MigrationObject
}


function Get-PreMigrationTaskStatus{
  param(
    [Parameter(Mandatory = $true)]$MigrationObject
  )

  $CheckPoints = @()


  #get definition
  $userProf = $USER_PROFILES | Where-Object {$_.licenseGroup -eq $MigrationObject.licenseGroup}
  $MigrationObject.UserProfile = $userProf.profileName
  
  #1. Check License
  $thisCheckPointLic = [PSCustomObject]@{
    Name = "Grant-License"
    isCompleted = $false
  }
  if(($MigrationObject.isO365Provioned -eq $true) -and ($userProf.licenseGroup -eq $MigrationObject.licenseGroup)){
    $thisCheckPointLic.isCompleted = $true
    
  }
  $CheckPoints += $thisCheckPointLic

  #2. Check SecProfile
  $thisCheckPointSec = [PSCustomObject]@{
    Name = "Assign-SecurityProfile"
    isCompleted = $false
  }
  if($userProf.securityProfileGroup -like "SU-EAST-SK-SK-SecurityProfile*"){
    $thisCheckPointSec.isCompleted = $true
  }
  $CheckPoints += $thisCheckPointSec

  #3. OneDrive
  $thisCheckPointSOD = [PSCustomObject]@{
    Name = "Prov-OneDrive"
    isCompleted = $false
  }
  if($MigrationObject.OneDriveStatus -eq "Active"){
    $thisCheckPointSOD.isCompleted = $true
  }
  $CheckPoints += $thisCheckPointSOD

  #4.Regional configuration
  $thisCheckPointSRC = [PSCustomObject]@{
    Name = "RegionalConfig"
    isCompleted = $false
  }
  if($MigrationObject.RegionalConfiguration -eq $true){
    $thisCheckPointSRC.isCompleted = $true
  }
  $CheckPoints += $thisCheckPointSRC

  #5. Account Enable
  $thisCheckPointSEnabled = [PSCustomObject]@{
    Name = "AccountEnabled"
    isCompleted = $false
  }
  if($MigrationObject.AccountEnabled -eq $true){
    $thisCheckPointSEnabled.isCompleted = $true
  }
  $CheckPoints += $thisCheckPointSEnabled

  #6. ODM Tasks
  $thisUserCollection = $MigrationObject.Wave
  $premigTasks = $__TASKTEMPLATES.templates | Where-Object {$_.stage -eq "preMigration"}
  #MailTasks
  $premigMailTask = $premigTasks | Where-Object {$_.type -eq "mailMigration"}
  
  $thisCheckPointSODMPremigMail = [PSCustomObject]@{
    Name = "ODM-PRE-MAIL"
    isCompleted = $false
  }
  if($MigrationObject.ODMMailInitialSyncTask -like "*$($premigMailTask.prefix)_$($thisUserCollection)_Mail*" ){
    $thisCheckPointSODMPremigMail.isCompleted = $true
  }
  $CheckPoints += $thisCheckPointSODMPremigMail

  #OneDriveTasks
  $thisCheckPointMPremigOD = [PSCustomObject]@{
    Name = "ODM-PRE-OD"
    isCompleted = $false
  }
  $premigODTask = $premigTasks | Where-Object {$_.type -eq "oneDriveMigration"}

  #Cases
  if($MigrationObject.ODMODInitialSyncTask -like "*$($premigODTask.prefix)_$($thisUserCollection)_One*"){
    $thisCheckPointMPremigOD.isCompleted = $true
  }elseif(($MigrationObject.ODMODInitialSyncTask -eq "NotMigrable") -and ($userProf.licenseGroup -eq $MigrationObject.licenseGroup)){
    $thisCheckPointMPremigOD.isCompleted = $true
  }
  $CheckPoints += $thisCheckPointMPremigOD

 
  #Join all checkpoints
  $CheckAll = $CheckPoints | Where-Object {$_.isCompleted -eq $false}

  if((($CheckAll | Measure-Object).Count) -gt 1){
    $MigrationObject.PreMigrationTaskStatus = "INCOMPLETE=$( $CheckAll.Name -join ';' )"
  }else{
    $MigrationObject.PreMigrationTaskStatus = "Completed"
  }

  return $MigrationObject

}


function Get-GlobalMigrationStatus {
  param(
    [Parameter(Mandatory = $true)]$MigrationObject
  )

  
  if(($MigrationObject.PreMigrationTaskStatus -eq "Completed") -and ($MigrationObject.AccountEnabled -ne $false)){
    $MigrationObject.GlobalMigrationStatus = "ReadyToMigrate"
  }

  if((($MigrationObject.isMailboxAccessDisabled -eq $true) -and ($MigrationObject.MailboxSwitchDirection -like "Switched"))){
    $MigrationObject.GlobalMigrationStatus = "Migrated"
  }

  return $MigrationObject

}

function Get-MigrationTaskStatus {
  
  param(
    [Parameter(Mandatory = $true)]$MigrationObject
  )

  if(([String]::IsNullOrEmpty($MigrationObject.ODMMailFinalSyncTask) -eq $false) -and ($MigrationObject.isMailboxAccessDisabled -eq $true) -and ($MigrationObject.isOneDriveReadOnlyEnabled -like "ReadOnly")){
    $MigrationObject.MigrationTaskStatus = "Completed"
  }

  return $MigrationObject
}


$DATA_IN = Import-Csv -Path $COMMON_FILE
$DATA_OUT = @()

$DATA_IN | ForEach-Object {
  #Capture object for best code understanding

  $ThisMigrationObject = $_
  Write-Host "Getting: $($_.TargetUserPrincipalName)"

  #Check DUA Requriements
  $ThisMigrationObject = Get-DUARequirement -MigrationObject $ThisMigrationObject
  #Check premigrationtasks
  $ThisMigrationObject = Get-PreMigrationTaskStatus -MigrationObject $ThisMigrationObject
  #Check migration tasks
  $ThisMigrationObject = Get-MigrationTaskStatus -MigrationObject $ThisMigrationObject
  #Check Global status
  $ThisMigrationObject = Get-GlobalMigrationStatus -MigrationObject $ThisMigrationObject



  #StoreObject
  $DATA_OUT += $ThisMigrationObject
}

$DATA_OUT | Export-Csv -Path $COMMON_FILE -NoTypeInformation
$DATA_OUT | Format-List