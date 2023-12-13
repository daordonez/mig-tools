<#
  Date: 10/01/2023
  Author: Diego Ordonez
  Synopsis : Collect OneDriveStatus from users in CSV File
#>

param(
  [Parameter(Mandatory=$false)][switch]$CreateConnection
)

$TODAY = Get-Date -Format yyyyMMddHHmm
$PATH_SCRIPT = $PSScriptRoot
$PATH_LOGS = "$($PATH_SCRIPT)\logs"
$PATH_CONFIG = "$($PATH_SCRIPT)\project_config.json"
$PATH_USERS_IN = "$($PATH_SCRIPT)\Users.csv"
$PATH_EXPORTS = "$($PATH_SCRIPT)\export"

$__CONFIG = Get-Content -Path $PATH_CONFIG | ConvertFrom-Json
$FILE_LOGS = $($__CONFIG.params.sharePoint.logsFileName)
$PATH_EXEC_LOG = "$($PATH_LOGS)\$($FILE_LOGS)"

#Logging
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

    Write-Error "LogFile '$($FILE_LOGS)' not found"
  }

}

#Create connection
$ADMIN_URL = $__CONFIG.params.sharePoint.adminUrl
Write-Log -Level INFO -LogMessage "Script Start" -VerboseOut

if($CreateConnection -eq $true){
  Write-Log -Level INFO -LogMessage "Connecting to: $($ADMIN_URL)" -VerboseOut
  Connect-SPOService -Url $ADMIN_URL
}



##load csv users
$usersIn = Import-Csv -Path $PATH_USERS_IN
$totalCount = $usersIn.count
Write-Log -Level INFO -LogMessage "Total users:$($totalCount) " -VerboseOut

$AllSitesOut = @()
Write-Log -Level INFO -LogMessage "Getting user information" -VerboseOut
$usersIn | ForEach-Object {

  Write-Host "Getting:$($_.SourceUserPrincipalName)"

  $thisUserSiteOb = [PSCustomObject]@{
    Owner = $_.SourceUserPrincipalName
    lockState = "NOTAPPLICABLE_F3"
  }

  $LicenseType = $_.LicenseType

  if($LicenseType -ne "F3"){

    ##Cast users to accepted format
    $clearAt = $_.SourceUserPrincipalName -replace '@', '_'
    $clearDot = $clearAt -replace '\.', '_'
  
    #Fill urls attributes
    #$thisUser.TransformedAddress = $clearDot
    $thisUserSite = "$($__CONFIG.params.sharePoint.baseODUrl)/$($clearDot)"
  
    $ThisFoundSite = Get-SPOSite -Identity $thisUserSite| Select-Object owner,lockstate

    $thisUserSiteOb.Owner = $ThisFoundSite.owner
    $thisUserSiteOb.lockState = $ThisFoundSite.lockstate
  }

  $AllSitesOut += $thisUserSiteOb

}

$AllSitesOut | Format-Table -AutoSize
