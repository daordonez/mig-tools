<#
  Date: 13/03/2023
  Author: Diego Ordonez
  Synopsis : Collect regional configuration from users in CSV File
#>

param(
  [Parameter(Mandatory=$false)][switch]$CreateConnection
)

#$TODAY = Get-Date -Format yyyyMMddHHmm
$PATH_SCRIPT = $PSScriptRoot
$PATH_LOGS = "$($PATH_SCRIPT)\logs"
$PATH_CONFIG = "$($PATH_SCRIPT)\project_config.json"
$PATH_COMMON = "$($PSScriptRoot)\common"
$PATH_USERS_IN = "$($PATH_COMMON)\MigrationGlobal.csv"
$COMMON_FILE = "$($PATH_COMMON)\MigrationGlobal.csv"
#$PATH_EXPORTS = "$($PATH_SCRIPT)\exports"

$__CONFIG = Get-Content -Path $PATH_CONFIG | ConvertFrom-Json
$FILE_LOGS = $($__CONFIG.params.exchangeOnline.logsFileName)
$PATH_EXEC_LOG = "$($PATH_LOGS)\$($FILE_LOGS)"


function Write-Log {

  param(
    [Parameter(Mandatory = $true)][ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]$Level,
    [Parameter(Mandatory = $true)]$LogMessage,
    [Parameter(Mandatory = $false)][switch]$VerboseOut
      
  )

  $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

  $Line = "$Stamp $Level $LogMessage"

  if (Test-Path -Path $PATH_EXEC_LOG) {

    Add-Content -Path $PATH_EXEC_LOG -Value $Line

    if ($VerboseOut -eq $true) {
      switch ($Level) {
        INFO {
          Write-Host $Line
        }
        WARN {
          Write-Host $Line -ForegroundColor Yellow
        }
        ERROR {
          Write-Host $Line -ForegroundColor Red
        }
        FATAL {
          Write-Host $Line -ForegroundColor Red
        }
        Default {
          Write-Host $Line
        }
      }
    }


  }
  else {

    Set-Content -Path $PATH_EXEC_LOG -Value ""
  }

}

function New-ConnectionSet {
  Write-Log -Level INFO -LogMessage "Creating SOURCE ($($__CONFIG.params.exchangeOnline.sourceTenantDomain)) connection" -VerboseOut
  Connect-ExchangeOnline -Prefix "$($__CONFIG.params.exchangeOnline.sourceConnectionPrefix)" -ShowBanner:$false

  Write-Log -Level INFO -LogMessage "Creating TARGET ($($__CONFIG.params.exchangeOnline.targetTenantDomain)) connection" -VerboseOut
  Connect-ExchangeOnline -Prefix "$($__CONFIG.params.exchangeOnline.targetConnectionPrefix)" -ShowBanner:$false
}

function New-UserMailboxes {
  
  param(
    [Parameter(Mandatory = $true)]$SourceUserPrincipalName,
    [Parameter(Mandatory = $true)]$TargetUserPrincipalName
  )

  $thisMailboxRegionalConf = [PSCustomObject]@{
    SourceUserPrincipalName                     = $SourceUserPrincipalName
    TargetUserPrincipalName                     = $TargetUserPrincipalName
    SourceDateFormat                            = ""
    SourceLanguage                              = ""
    SourceDefaultFolderNameMatchingUserLanguage = ""
    SourceTimeFormat                            = ""
    SourceTimeZone                              = ""
    TargetDateFormat                            = ""
    TargetLanguage                              = ""
    TargetDefaultFolderNameMatchingUserLanguage = ""
    TargetTimeFormat                            = ""
    TargetTimeZone                              = ""
    ConfigurationCopied                         = $false
  }

  return $thisMailboxRegionalConf
}


function Get-RegionalConfig {
  param(
    [Parameter(Position = 0, Mandatory = $true)]$UserPrincipalname,
    [Parameter(Mandatory = $true)][ValidateSet("SRC", "TRG")]$Environment
  )

  $RegionalConfig

  switch ($Environment) {
    SRC { 
      $RegionalConfig = Get-SRCMailboxRegionalConfiguration -Identity $UserPrincipalname
    }
    TRG {
      $RegionalConfig = Get-TRGMailboxRegionalConfiguration -Identity $UserPrincipalname
    }
  }

  return $RegionalConfig
}

function Get-ConfigurationCopyStatus {
  param(
    [Parameter(Mandatory = $true)]$RegionalConfigurationObject
  )

  $isCopied = $false
  $copyScore = 0

  if($RegionalConfigurationObject.SourceDateFormat -eq $RegionalConfigurationObject.TargetDateFormat){
    $copyScore++
  }

  if($RegionalConfigurationObject.SourceLanguage -eq $RegionalConfigurationObject.TargetLanguage){
    $copyScore++
  }
  if($RegionalConfigurationObject.SourceDefaultFolderNameMatchingUserLanguage -eq $RegionalConfigurationObject.TargetDefaultFolderNameMatchingUserLanguage){
    $copyScore++
  }
  if($RegionalConfigurationObject.SourceTimeFormat -eq $RegionalConfigurationObject.TargetTimeFormat){
    $copyScore++
  }
  if($RegionalConfigurationObject.SourceTimeZone -eq $RegionalConfigurationObject.TargetTimeZone){
    $copyScore++
  }

  if($copyScore -ge 3){
    $isCopied = $true
  }

  return $isCopied
}



if($CreateConnection -eq $true){
  Disconnect-ExchangeOnline -Confirm:$false
  New-ConnectionSet
}

Write-Log -Level INFO -LogMessage "Loading users" -VerboseOut
$UsersConfTemp = @()
$UsersDumpRegionalConf =  @()
$UsersMBxRegionalConf = Import-Csv -Path $PATH_USERS_IN

Write-Log -Level INFO -LogMessage "Gathering source information" -VerboseOut
#1. Source
$UsersMBxRegionalConf | ForEach-Object {

  $thisNewSource = New-UserMailboxes -SourceUserPrincipalName $_.SourceUserPrincipalname -TargetUserPrincipalName $_.TargetUserPrincipalName

  $thisUserRegionalConf = Get-RegionalConfig -UserPrincipalname $thisNewSource.SourceUserPrincipalName -Environment SRC

  #Mapping values
  $thisNewSource.SourceDateFormat = $thisUserRegionalConf.DateFormat
  $thisNewSource.SourceLanguage = $thisUserRegionalConf.Language
  $thisNewSource.SourceDefaultFolderNameMatchingUserLanguage = $thisUserRegionalConf.DefaultFolderNameMatchingUserLanguage
  $thisNewSource.SourceTimeFormat = $thisUserRegionalConf.TimeFormat
  $thisNewSource.SourceTimeZone = $thisUserRegionalConf.TimeZone

  $UsersConfTemp += $thisNewSource

}


#2. Target
Write-Log -Level INFO -LogMessage "Gathering target information" -VerboseOut
$UsersConfTemp | ForEach-Object {

  if($_.TargetUserPrincipalName -notlike "0_NOMATCH"){

    $thisUserRegionalConfTarget = Get-RegionalConfig -UserPrincipalname $_.TargetUserPrincipalName -Environment TRG
  
    $_.TargetDateFormat = $thisUserRegionalConfTarget.DateFormat
    $_.TargetLanguage = $thisUserRegionalConfTarget.Language
    $_.TargetDefaultFolderNameMatchingUserLanguage = $thisUserRegionalConfTarget.DefaultFolderNameMatchingUserLanguage
    $_.TargetTimeFormat = $thisUserRegionalConfTarget.TimeFormat
    $_.TargetTimeZone = $thisUserRegionalConfTarget.TimeZone
  }


  $UsersDumpRegionalConf += $_
}

Write-Log -Level INFO -LogMessage "Comparing resources" -VerboseOut
$UsersDumpRegionalConf | ForEach-Object{

  $_.ConfigurationCopied = Get-ConfigurationCopyStatus -RegionalConfigurationObject $_

}

$UsersMBxRegionalConf | ForEach-Object{

  $userDumpableObject = $_

  $thisUserExoRegionalConf = $UsersDumpRegionalConf | Where-Object {($_.SourceUserPrincipalName -eq $userDumpableObject.SourceUserPrincipalName) -and ($_.TargetUserPrincipalname -eq $userDumpableObject.TargetUserPrincipalName)}

  $_.RegionalConfiguration = $thisUserExoRegionalConf.ConfigurationCopied
}

Write-Log -Level INFO -LogMessage "Dumping results into 'Common'" -VerboseOut
$UsersMBxRegionalConf | Export-Csv -Path $COMMON_FILE -NoTypeInformation
