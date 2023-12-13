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
$PATH_USERS_IN = "$($PATH_SCRIPT)\common\MigrationGlobal.csv"
$PATH_COMMON = "$($PSScriptRoot)\common"
$COMMON_FILE = "$($PATH_COMMON)\MigrationGlobal.csv"
$PATH_EXPORTS = "$($PATH_SCRIPT)\exports"

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


$usersTransformed = @()

Write-Log -Level INFO -LogMessage "Getting user information" -VerboseOut
$usersIn | ForEach-Object {

  $thisUser = [PSCustomObject]@{
    UserPrincipalName = $_.TargetUserPrincipalname
    DisplayName = ""
    TransformedAddress = ""
    UserOneDriveSite = ""
    Status = ""
    OneDriveOwner = ""
  }

  ##Cast users to accepted format
  $clearAt = $thisUser.UserPrincipalName -replace '@', '_'
  $clearDot = $clearAt -replace '\.', '_'

  #Fill urls attributes
  $thisUser.TransformedAddress = $clearDot
  $thisUser.UserOneDriveSite = "$($__CONFIG.params.sharePoint.baseODUrl)/$($clearDot)"

  #Get user OneDrive

  #Write-host "`tGetting:$($thisUser.UserPrincipalName)" -NoNewline

  try {
    $oneDriveStatus = Get-SPOSite -Identity $thisUser.UserOneDriveSite -ErrorAction SilentlyContinue
    #Write-Host "`t OK Active" -ForegroundColor Green
  }
  catch {
    #Write-Host "`t ! Pending" -ForegroundColor Yellow
  }


  if ($null -ne $oneDriveStatus) {
    $thisUser.DisplayName =  $oneDriveStatus.Title
    $thisUser.Status = $oneDriveStatus.Status
    $thisUser.OneDriveOwner = $oneDriveStatus.Owner
  }else{
    $thisUser.Status = "Not provisioned"
  }


  #StoreCurrent data in global
  $_.OneDriveStatus = $thisUser.Status

  $usersTransformed += $thisUser
  $oneDriveStatus = $null

}

$totalActive = ($usersTransformed | Where-Object {$_.Status -like "Active"} | measure).Count
$pending = $totalCount - $totalActive
Write-Log -Level INFO -LogMessage "Summary: Total:$($totalCount),Active:,$($totalActive),Pending:$($pending)" -VerboseOut

Write-Log -Level INFO -LogMessage "Exporting results. PATH:$($PATH_EXPORTS)" -VerboseOut
$usersIn | Export-Csv -Path $COMMON_FILE -NoTypeInformation
$usersTransformed | Export-Csv -Path "$($PATH_EXPORTS)\$($TODAY)_OneDriveStatus.csv" -NoTypeInformation

#$usersTransformed | Format-Table  DisplayName,UserPrincipalName,Status,UserOneDriveSite -AutoSize

$usersIn | Format-List

