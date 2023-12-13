<#
  Date: 17/01/2023
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
$PATH_COMMON = "$($PSScriptRoot)\common"
$COMMON_FILE = "$($PATH_COMMON)\MigrationGlobal.csv"
$PATH_EXPORT = "$($PSScriptRoot)\exports"
$__CONFIG = Get-Content -Path $PATH_CONFIG | ConvertFrom-Json
$TODAY =  Get-Date

#Check for createConnection
if($CreateConnection -eq $true){
  Write-Host "Loading Msol Module"
  Import-Module -Name MSOnline
  write-Host "Creating connection"
  Connect-MsolService
}


$UsersIn = Import-Csv -Path $COMMON_FILE


$UsersOut = @()
$UsersIn | ForEach-Object{

  $isO365Provioned = $false

  if($_.UserMailboxType -ne "SharedMailbox"){

    #Write-Host "Getting: $($_.TargetUserPrincipalName)"
  
    $thisUser = Get-MsolUser -UserPrincipalName $_.TargetUserPrincipalName -ErrorAction SilentlyContinue
  
    if($null -ne $thisUser.Licenses.AccountSkuId){
      $licenseDetail = (($thisUser.Licenses.AccountSkuId) -replace 'reseller-account:','') -join ';'
      $isO365Provioned = $true
    }else{
      $licenseDetail = ""
    }
  }else{
    #Write-Host "Skipping: $($_.SourceUserPrincipalName).MailboxType=$($_.UserMailboxType)"
    $licenseDetail = "SHAREDMBX_NOTNEEDED"
  }


  

  $userDetails = [PSCustomObject]@{
    UserPrincipalName = $_.UserPrincipalName
    DisplayName = $thisUser.DisplayName
    isLicensed = $thisUser.isLicensed
    isO365Provioned = $isO365Provioned
    LicenseDetail = $licenseDetail
  }

  $_.LicenseDetail = $userDetails.LicenseDetail
  $_.isO365Provioned = $userDetails.isO365Provioned
  
  $UsersOut += $userDetails

}


$usersIn | Export-Csv -Path $COMMON_FILE -NoTypeInformation
$todayToString = $TODAY | Get-Date -Format "yyyyMMddHHmm"
$UsersOut | Export-Csv -Path "$($PATH_EXPORT)\$($todayToString)_MSOLLicenseStatusDetail.csv" -NoTypeInformation

$UsersIn | Format-List