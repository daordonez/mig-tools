<#
  Date: 2023-04-12
  Environment: PRO-SK-Onpremise
  Synopsis: This script will remove the domain passed in param '-ToRemoveDomain' from the current local Active Directory ,also,
            it will set up the new UPN with the given domain by '-ReplacementDomain' params.
#>

param(
  [Parameter(Mandatory=$true)]$ToRemoveDomain,
  [Parameter(Mandatory=$true)]$ReplacementDomain
)

#Const
$PATH_ROOT = $PSScriptRoot
$TODAY = Get-Date -Format "yyyyMMddHHmm"
$PATH_EXEC_LOG = "$($PATH_ROOT)\log\SetUPN.log"
$PATH_OUTPUT = "$($PATH_ROOT)\output"
$PATH_USERSIN = "$($PATH_ROOT)\Users.csv"

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

    Write-Error "LogFile 'SetUPN.log' not found"
  }

}

#Runtime
Write-Log -Level INFO -LogMessage "Script SetUserPrincipalName start" -VerboseOut
Write-Log -Level INFO -LogMessage "Input Params: ToRemoveDomain=$($ToRemoveDomain). ReplacementDomain=$($ReplacementDomain)" -VerboseOut
Write-Log -Level INFO -LogMessage "Loading users from CSV" -VerboseOut

$UsersIN = Import-Csv -Path $PATH_USERSIN

$TotalUsersIn = ($UsersIN | Measure-Object).Count
Write-Log -Level INFO -LogMessage "Users to modify.TotalCount=$($TotalUsersIn)" -VerboseOut
Write-Log -Level INFO -LogMessage "Starting deployment" -VerboseOut

$ObjectsDeployment = @()
$UsersIN | ForEach-Object{
  $thisUserPrincipalName = $_.UserPrincipalName

  $ThisUPNChange = [PSCustomObject]@{
    OldUPN = $thisUserPrincipalName
    NewUPN = ""
    CurrentUPN = ""
    Completed = $false
  }

  $NewUPNprefix = $thisUserPrincipalName.split('@')[0]
  $newUPN = "$($NewUPNprefix)@$($ReplacementDomain)"

  $ThisUPNChange.NewUPN = $newUPN

  Set-MsolUserPrincipalName -UserPrincipalName $thisUserPrincipalName -NewUserPrincipalName $newUPN

  $ObjectsDeployment += $ThisUPNChange

}

Write-Log -Level INFO -LogMessage "Getting current 'UserPrincipalName'" -VerboseOut

$ObjectsDeployment | ForEach-Object {
  $thisUPNToCheck = $_.NewUPN

  $UPNChanged = Get-MsolUser -UserPrincipalName $thisUPNToCheck -ErrorAction SilentlyContinue

  if($null -ne $UPNChanged){
    $_.CurrentUPN = $UPNChanged.UserPrincipalName

    if($UPNChanged.UserPrincipalName -like "*$($ReplacementDomain)"){
      $_.Completed = $true
    }

  }
}

Write-Log -Level INFO -LogMessage "Output results in UI" -VerboseOut
$parseToRemoveDomain = $ToRemoveDomain -replace '\.','_'
$PathExportOutput = "$($PATH_OUTPUT)\$($TODAY)_SetUserPrincipalName_$($parseToRemoveDomain).csv"
Write-Log -Level INFO -LogMessage "Output results exported. PATH=$($PathExportOutput)" -VerboseOut
$ObjectsDeployment | Export-Csv -Path $PathExportOutput -NoTypeInformation
$ObjectsDeployment | Out-GridView
Write-Log -Level INFO -LogMessage "Script end" -VerboseOut






