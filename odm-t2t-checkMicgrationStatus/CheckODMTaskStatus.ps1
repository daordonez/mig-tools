<#
  Date: 10/01/2023
  Author: Diego Ordonez
  Synopsis : Collect ODM from users MigrationGlobal.csv
#>

param(
  [Parameter(Mandatory = $false)][switch]$CreateConnection
)

$TODAY = Get-Date -Format yyyyMMddHHmm
#$PATH_SCRIPT = $PSScriptRoot
$PATH_SCRIPT = "."
$PATH_LOGS = "$($PATH_SCRIPT)\logs"
$PATH_COMMON = "$($PSScriptRoot)\common"
$PATH_CONFIG = "$($PATH_SCRIPT)\project_config.json"
$PATH_ODMEVENTS = "$($PATH_COMMON)\ODMGLobalEvents.csv"
$PATH_USERS_IN = "$($PATH_COMMON)\MigrationGlobal.csv"
$COMMON_FILE = "$($PATH_COMMON)\MigrationGlobal.csv"
$PATH_EXPORTS = "$($PATH_SCRIPT)\exports"

$__CONFIG = Get-Content -Path $PATH_CONFIG | ConvertFrom-Json
$FILE_LOGS = $($__CONFIG.params.sharePoint.logsFileName)
$PATH_EXEC_LOG = "$($PATH_LOGS)\$($FILE_LOGS)"


function Get-SwitchMigrationTask {
  param(
    [Parameter(Mandatory = $true)]$MigrationObject
  )

  #Get all de switch tasks
  $allTasks = Get-OdmTask | Where-Object { $_.Type -like "Mailbox Switch" }

  $allTasks = $allTasks | Where-Object { $_ -like "*Mail_Forwarding*" }

  $allTasks | ForEach-Object {
    $usersInTask = $_ | Get-OdmObject | Select-Object TargetUserPrincipalname

    #If task contains user
    if ($usersInTask.TargetUserPrincipalName.contains($($MigrationObject.TargetUserPrincipalname))) {
      $ThisUserFWDTask = [PSCustomObject]@{
        SwitchTaskname  = $_.Name
        SwitchDirection = $_.SwitchDirection
      }

      $MigrationObject.ODMMailboxSwitchTask = $ThisUserFWDTask.SwitchTaskname
      $MigrationObject.MailboxSwitchDirection = $ThisUserFWDTask.SwitchDirection
       
    }
  }
  return $MigrationObject
}
#Updating Events
Write-Host "Pulling ODM Events to local"
$ODMEVENTS_IN = Import-Csv -Path $PATH_ODMEVENTS
$DATA_IN = Import-Csv -Path $PATH_USERS_IN
$DATA_OUT = @()



#Store this collection
$AllWavesInfo = @()
Write-Host "Getting waves information"
$WaveNames = ($DATA_IN | Group-Object Wave).Name

Write-Host "Getting Waves: " -NoNewline

$WaveNames = $WaveNames | Where-Object { $_ -ne $null }
$WaveNames | ForEach-Object {


  Write-Host "$($_)," -NoNewline
  $collectionName = @{'name' = $_ }
  $ThisWave = Get-OdmCollection -WildcardFilter $collectionName
  
  $RawODMData = $ThisWave | Get-OdmObject

  $thisWaveJoin = [PsCustomObject]@{
    WaveId = $_
    ODMID = $ThisWave._Id
    UsersIN = $RawODMData.TargetUserPrincipalName
    RAWODMData = $RawODMData
  }

  $AllWavesInfo += $thisWaveJoin
}
Write-Host

$AllWavesInfo = $AllWavesInfo | Where-Object { $_.WaveId -ne $null }

#Arrange events
Write-Host "Grouping tasks"
Write-Host "`t > Mail"
$MailMigartionTasks = ($ODMEVENTS_IN | Group-Object Category | Where-Object { $_.Name -eq "Mail Migration" }).group | Group-Object Taskname
Write-Host "`t > Mail Switch"
$MailSwitchTasks = ($ODMEVENTS_IN | Group-Object Category | Where-Object { $_.Name -eq "Mailbox Switch" }).group | Group-Object Taskname
Write-Host "`t > OneDrive"
$ODMigartionTasks = ($ODMEVENTS_IN | Group-Object Category | Where-Object { $_.Name -eq "OneDrive Migration" }).group | Group-Object Taskname
Write-Host "`t > DUA"
$DUASwitchTasks = ($ODMEVENTS_IN | Group-Object Category | Where-Object { $_.Name -eq "Desktop Agent" }).group | Group-Object Taskname

$DATA_IN | ForEach-Object {

  #UserWave
  $thisUserWave = $_.Wave

  if ($_.UserMailboxType -ne "SharedMailbox") {

    $searchString = @{'targetUserPrincipalName' = $_.TargetUserPrincipalname }
    $thisUser = Get-OdmObject -WildcardFilter $searchString

    
    
  }
  else {
      $ThisODMWave = $AllWavesInfo | Where-Object { $_.WaveId -eq $thisUserWave }

      $SourceUPN = $_.SourceUserPrincipalName

      #Check is in wave
      $isSourceODWave = $ThisODMWave.RAWODMData | Where-Object {$_.SourceUserPrincipalName -eq $SourceUPN}

      $thisUser = $isSourceODWave 

  }

  

  #Check if user is in collection

  if ($null -ne $thisUserWave) {
    $isInODMWave = ($AllWavesInfo | Where-Object { $_.WaveId -eq $thisUserWave }).UsersIN.contains($_.TargetUserPrincipalname)
    if ($isInODMWave) {
      $_.ODMCollectionName = $thisUserWave
    }else{

      if($null -ne $isSourceODWave){
        $_.ODMCollectionName = $thisUserWave
      }else{
        $_.ODMCollectionName = "NOTCOLLECTION"
      }

      $thisUser = $isSourceODWave
    }
    

  }

  

  

  #Filter Mail Migration Tasks
  #Pre Migration Tasks
  $thisuserTasks_MAIL = $MailMigartionTasks | Where-Object { $_.Group.ObjectId -eq $thisUser._Id }

  $PremigTaskName = ($thisuserTasks_MAIL | Where-Object { $_.Name -like "*T10_$($thisUserWave)_Mail*"}).Name -join ';'
  $_.ODMMailInitialSyncTask = $PremigTaskName

  #Get timestamp
  if(([string]::IsNullOrEmpty($_.ODMMailInitialSyncTask)) -eq $false){
    

    $_.PreMigrationTasksTimeStamp = (($thisuserTasks_MAIL | Where-Object {$_.Name -eq $PremigTaskName}).Group | Sort-Object Timestamp)[0].Timestamp

  }

  #Migration Tasks
  $MigtaskName = ($thisuserTasks_MAIL | Where-Object { $_.Name -like "*T1_$($thisUserWave)_Mail*"}).Name -join ';'
  $_.ODMMailFinalSyncTask = $MigtaskName

  #Get timestamp
  if(([string]::IsNullOrEmpty($MigtaskName)) -eq $false){
    

    $_.MigrationTasksTimeStamp = (($thisuserTasks_MAIL | Where-Object {$_.Name -eq $MigtaskName}).Group | Sort-Object TimeStamp)[0].Timestamp

  }

  #Mail switch task
  $thisuserTasks_MAILFWD = $MailSwitchTasks | Where-Object { $_.Group.ObjectId -eq $thisUser._Id }
  $_.ODMMailboxSwitchTask = $thisuserTasks_MAILFWD | Where-Object { ($_.Name -like "*Mailbox_Switch*") -or ($_.Name -like "*MailboxSwitch*") }


  #Filter One Drive Migration tasks
  #Pre Migration Tasks
  $thisuserTasks_ONEDRIVE = $ODMigartionTasks | Where-Object { $_.Group.ObjectId -eq $thisUser._Id }
  $_.ODMODInitialSyncTask = (($thisuserTasks_ONEDRIVE | Where-Object { $_.Name -like "*T10_$($thisUserWave)_OneDrive_Migration*" -or ($_.Name -like "*T10_$($thisUserWave)_OneDriveMigration*") })).Name -join ';'
  #Migration Tasks
  $_.ODMODFinalSyncTask = (($thisuserTasks_ONEDRIVE | Where-Object { $_.Name -like "*T1_$($thisUserWave)_OneDrive_Migration*" -or ($_.Name -like "*T1_$($thisUserWave)_OneDriveMigration*") })).Name -join ';'

  #Check if it is OneDrive migrable
  if(([string]::IsNullOrEmpty($_.ODMODInitialSyncTask)) -eq $true){
    #by operator decition
    if($null -eq $thisuserTasks_ONEDRIVE){
      $_.ODMODInitialSyncTask = "NotMigrable"
      $_.ODMODFinalSyncTask = "NotMigrable"
      #By ODM event
    }elseif(($thisuserTasks_ONEDRIVE.group | Where-Object {$_.ObjectId -eq $thisUser._Id}).Message.contains("OneDrive does not exist in the source tenant. Nothing to migrate.")){
      $_.ODMODInitialSyncTask = "NotMigrable"
      $_.ODMODFinalSyncTask = "NotMigrable"
    }
  }
  
  #Filter Application migration task
  $thisuserTasks_DUA = $DUASwitchTasks | Where-Object { $_.Group.ObjectId -eq $thisUser._Id }
  $_.ODMAppSwitchTask = ($thisuserTasks_DUA | Where-Object { ($_.Name -like "*T1_$($thisUserWave)_SwitchApplications*") -or ($_.Name -like "*T1_$($thisUserWave)_Switch_Applications*") }).Name -join ';'

  #Fill user ODM object
  $_.ODMMailboxSwitchTask = $thisUser.SwitchStatus
  $_.MailboxSwitchDirection = $thisUser.SwitchDirection
  $_.MailMigrationRate = $thisUser.MailMigrationProgress
  $_.OneDriveMigrationRate = $thisUser.OneDriveMigrationProgress

 
  #Populate new account for SharedMailbox
  if($_.UserMailboxType -eq "SharedMailbox"){
    $_.TargetUserPrincipalname = $thisUser.TargetUserPrincipalName
  }

  $finalObject = $_

  $finalObject  

  $DATA_OUT += $finalObject
}


$DATA_OUT | Export-Csv -Path $COMMON_FILE -NoTypeInformation
$DATA_OUT | Export-Csv -Path "$($PATH_EXPORTS)\$($TODAY)_ODMTaskStatus.csv" -NoTypeInformation
$DATA_OUT | Format-List