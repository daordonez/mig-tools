<#
  Date: 16/01/2023
  Author: Diego Ordonez
  Synopsis: Check user migration status on basis of 'user_profiles' definitions
#>

param(
  [Parameter(Position = 0, Mandatory = $false)]$TargetUserPrincipalname,
  [Parameter(Position = 1, Mandatory = $false)]$SourceUserPrincipalname,
  [Parameter(Position = 1, Mandatory = $false)][switch]$CreateConnection

)

$PATH_SCRIPT = $PSScriptRoot
$PATH_CONFIG = "$($PATH_SCRIPT)\project_config.json"
$PATH_USERPROFILES = "$($PATH_SCRIPT)\user_profiles.json"
$PATH_USERS_IN = "$($PSScriptRoot)\Users.csv"
$PATH_COMMON = "$($PSScriptRoot)\common"
$COMMON_FILE = "$($PATH_COMMON)\MigrationGlobal.csv"
$REPOSITORY_FILE = "$($PATH_COMMON)\LocalRepository_ScheduledUsers.csv"
$PATH_EXPORT = "$($PSScriptRoot)\exports"
$__CONFIG = Get-Content -Path $PATH_CONFIG | ConvertFrom-Json
$__USERPROFILES = Get-Content -Path $PATH_USERPROFILES | ConvertFrom-Json
$TODAY = Get-Date

#UserProfilesParams
$userCommonParams = $__USERPROFILES.params.commonAttributes
$licenseGroupPrefix = $userCommonParams.licenseGroupPrefix
$securityGroupPrefix = $userCommonParams.securityGroupPrefix

#LocalRepositoryParams
$REPO_PARAMS = Import-Csv -Path $REPOSITORY_FILE


#Check for createConnection
if ($CreateConnection -eq $true) {
  Write-Host "Loading AzureAD Module"
  Import-Module -Name AzureAD
  write-Host "Creating connection"
  Connect-AzureAD -TenantId $__CONFIG.params.o365TenantConfiguration.AADTenantId
}



function New-UserMigrationObject {
  param(
    [Parameter(Position = 0, Mandatory = $true)]$TargetUserPrincipalname,
    [Parameter(Position = 1, Mandatory = $true)]$SourceUserPrincipalname
  )

  $thisUserMigrationObject = [PSCustomObject]@{
    DisplayName                   = ""
    SourceUserPrincipalName       = $SourceUserPrincipalname
    TargetUserPrincipalname       = $TargetUserPrincipalname
    UserProfile = ""
    AccountEnabled                = ""
    Country                       = ""
    LicenseDetail                 = ""
    hasLoggedIn                   = $false
    LastRefreshMigrationInventory = ""
    RequiresDUADeployment         = $false
    SecurityProfileGroup          = ""
    LicenseGroup                  = ""
    OneDriveStatus                = ""
    RegionalConfiguration         = $false
    isO365Provioned               = $false
    isPasswordSet                 = $false
    PasswordSetTimeStamp          = ""
    UserMailboxType               = ""
    Wave                          = ""
    WaveClosingTimeStamp          = ""
    ODMCollectionName             = ""
    ODMMailInitialSyncTask        = ""
    ODMODInitialSyncTask          = ""
    ODMMailboxSwitchTask          = ""
    ODMAppSwitchTask              = ""
    ODMMailFinalSyncTask          = ""
    ODMODFinalSyncTask            = ""
    PreMigrationTasksTimeStamp    = ""
    PreMigrationTaskStatus        = ""
    DataMigrationReview           = ""
    CommunicationPhase            = 0
    isMailboxAccessDisabled       = ""
    MailboxAccessDisableTimeStamp = ""
    MailboxSwitchDirection        = ""
    isOneDriveReadOnlyEnabled     = ""
    OneDrivereadOnlyTimeStamp     = ""
    MigrationTasksTimeStamp       = ""
    MigrationTaskStatus           = ""
    MailMigrationRate             = 0
    OneDriveMigrationRate         = 0
    GlobalMigrationStatus         = "NotMigrated"
    
  }

  return $thisUserMigrationObject
}


function Get-AADUserParams {
  param(
    [Parameter(Mandatory = $true)]$MigrationObject
  )

  #Check 'NO_MATCH'
  if ($MigrationObject.TargetUserPrincipalname -notlike "0_NOMATCH") {
    
    $thisAADUser = Get-AzureADUser -ObjectId $MigrationObject.TargetUserPrincipalname -ErrorAction SilentlyContinue
    $thisAADMemberships = Get-AzureADUserMembership -ObjectId $MigrationObject.TargetUserPrincipalname -ErrorAction SilentlyContinue
  
    #DisplayName
    $MigrationObject.DisplayName = $thisAADUser.DisplayName
    #Country
    $MigrationObject.Country = $thisAADUser.Country
    #is Account Enabled
    $MigrationObject.AccountEnabled = $thisAADUser.AccountEnabled
    #LicenseGroup
    $MigrationObject.LicenseGroup = ($thisAADMemberships | Where-Object { $_.Displayname -like "$($licenseGroupPrefix)*" }).DisplayName
    #SecurityProfile
    $MigrationObject.SecurityProfileGroup = ($thisAADMemberships | Where-Object { $_.Displayname -like "$($securityGroupPrefix)*" }).DisplayName
  }


  return $MigrationObject
}

function Get-LocalRepositoryParams {
  param(
    [Parameter(Mandatory = $true)]$MigrationObject
  )

  #1. Check if user is scheduled
  if ($REPO_PARAMS.TargetUserPrincipalName.contains($MigrationObject.TargetUserPrincipalname)) {

    if ($MigrationObject.TargetUserPrincipalName -like "0_NOMATCH") {

      #Find by source UPN
      $thisUserInInventory = $REPO_PARAMS | where-object { $_.SourceUserPrincipalName -eq $MigrationObject.SourceUserPrincipalname }
      $MigrationObject.UserMailboxType = $thisUserInInventory.MailboxType
      $MigrationObject.Wave = $thisUserInInventory.Wave
      $MigrationObject.WaveClosingTimeStamp = $thisUserInInventory.WaveClosingTimeStamp
      $MigrationObject.LastRefreshMigrationInventory = $thisUserInInventory.LastRefreshTimeStamp

      if($thisUserInInventory.MailboxType -eq "SharedMailbox"){
        $MigrationObject.isMailboxAccessDisabled = "SHAREDMBX_NOTNEEDED"
        $MigrationObject.isOneDriveReadOnlyEnabled = "SHAREDMBX_NOTNEEDED"
        $MigrationObject.MailboxAccessDisableTimeStamp = "SHAREDMBX_NOTNEEDED"
        $MigrationObject.OneDrivereadOnlyTimeStamp = "SHAREDMBX_NOTNEEDED"
      }

    }
    else {
      #Find by target UPN (Default)
      $thisUserInInventory = $REPO_PARAMS | where-object { $_.TargetUserPrincipalName -eq $MigrationObject.TargetUserPrincipalname }
      $MigrationObject.UserMailboxType = $thisUserInInventory.MailboxType
      $MigrationObject.Wave = $thisUserInInventory.Wave
      $MigrationObject.WaveClosingTimeStamp = $thisUserInInventory.WaveClosingTimeStamp
      $MigrationObject.LastRefreshMigrationInventory = $thisUserInInventory.LastRefreshTimeStamp
      $MigrationObject.isMailboxAccessDisabled = $thisUserInInventory.isMailboxDisabled
      $MigrationObject.isOneDriveReadOnlyEnabled = $thisUserInInventory.isOneDriveReadOnly
      $MigrationObject.MailboxAccessDisableTimeStamp = $thisUserInInventory.DisableMailboxTimeStamp
      $MigrationObject.OneDrivereadOnlyTimeStamp = $thisUserInInventory.OneDriveReadOnlyTimeStamp
    }


  }
  else {
    $MigrationObject.Wave = "NOT_WAVE_FOUND"
  }

  return $MigrationObject

}




$OutUsers = @()


Import-CSV -Path $PATH_USERS_IN | ForEach-Object {

  $ThisMigrationUser = New-UserMigrationObject -TargetUserPrincipalname $_.TargetUserPrincipalname -SourceUserPrincipalname $_.SourceUserPrincipalname
  
  $ThisMigrationUser = Get-AADUserParams -MigrationObject $ThisMigrationUser

  $ThisMigrationUser = Get-LocalRepositoryParams -MigrationObject $ThisMigrationUser

  $OutUsers += $ThisMigrationUser
}

$OutUsers | Format-List 
$todayToString = $TODAY | Get-Date -Format "yyyyMMddHHmm"
$thisExportableName = "$($PATH_EXPORT)\$($todayToString)_AADUserStatus.csv"
$OutUsers | Export-Csv  -Path $thisExportableName -NoTypeInformation


<#
$CurrentData = Import-Csv -Path $COMMON_FILE

$semiFinalData = $CurrentData + $OutUsers
$finalData = $semiFinalData | Sort-Object TargetUserPrincipalname
#>
$OutUsers | Export-Csv -Path $COMMON_FILE -NoTypeInformation
