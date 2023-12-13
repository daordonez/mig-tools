<#
  Date : 01/01/2023
  Environment: Production
  Synopsis: Set personal OneDrive in 'Read-Only' mode
#>

param(
  [Parameter(Mandatory = $false)][switch]$CreateConnection,
  [Parameter(Mandatory = $true)]$Collection,
  [Parameter(Mandatory = $false)]
  [ValidateSet("Enable", "Disable")]
  $Operation = "Enable"
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

#Generate RepoPATH
$CurrentUser = (whoami).split("\\")[1]
$REPO_PATH = Join-Path -Path $__CONFIG.params.baseRepoPathPrefix -ChildPath "$($CurrentUser)\$($__CONFIG.params.baseReportPathSufix)"

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

function Set-UserSites {
  
  param(
    [Parameter(Mandatory = $true)]$UserList,
    [Parameter(Mandatory = $true)][ValidateSet("ReadOnly", "Unlock")]$LockState
  )

  #Return Value
  $usersOut = @()

  #Proc
  $usersIn | ForEach-Object {

    #Generate Object
    $thisUserSet = [PSCustomObject]@{

      ConfigurationTimeStamp = Get-Date
      UserPrincipalname      = $_.SourceUserPrincipalName
      LicenseType            = $_.LicenseType
      PersonalSite           = ""
      PostConfigurationState = ""
    }

    #Generate OneDrive for each user
    $clearDot = $($thisUserSet.UserPrincipalname) -replace '\.', '_'
    $clearAt = $clearDot -replace '@', '_'
    $thisClearString = $clearAt

    #Concat to OD base url
    $BaseUrl = $__CONFIG.params.sharePoint.baseODUrl
    $thisUserSet.PersonalSite = "$($BaseUrl)/$($thisClearString)"
    Write-Host "Setting: $($thisUserSet.UserPrincipalname)"

    #Check license Type
    $LicenseType = $thisUserSet.LicenseType
    if ($LicenseType -ne "F3"){

      #Check if it is already configured
      $preConf = (Get-SPOSite -Identity $thisUserSet.PersonalSite).LockState
  
      if ($preConf -notlike $lockState) {
      
        Write-Host "`t - PersonalSite: $($thisUserSet.PersonalSite)"
  
        #Setting site
        Set-SPOSite -Identity $thisUserSet.PersonalSite -LockState $LockState
  
        Start-Sleep -Seconds 3 
        Write-Host "`t - LockState: " -NoNewline
  
        #Check conf status
        $postConfStatus = (Get-SPOSite -Identity $thisUserSet.PersonalSite).LockState
  
        if ($postConfStatus -eq $lockState) {
          Write-Host "OK" -ForegroundColor Green
        }
        else {
          Write-Host "!" -ForegroundColor Yellow
        }
        #Saving current status
        $thisUserSet.PostConfigurationState = $postConfStatus
      }
      else {
        Write-Host "`t - Already done. Skipping site" -ForegroundColor Yellow
        $thisUserSet.PostConfigurationState = $preConf
      }
    }else{
      #Saving current status
      $thisUserSet.PostConfigurationState = "NOTAPPLICABLE"
      Write-Host "`t - Not OneDrive (F3 license). Skipping site" -ForegroundColor Yellow
    }

    $usersOut += $thisUserSet
  }

  return $usersOut

}


#Create connection
$ADMIN_URL = $__CONFIG.params.sharePoint.adminUrl
Write-Log -Level INFO -LogMessage "Script Start" -VerboseOut

if ($CreateConnection -eq $true) {
  Write-Log -Level INFO -LogMessage "Connecting to: $($ADMIN_URL)" -VerboseOut
  Connect-SPOService -Url $ADMIN_URL
}

Write-Log -Level INFO -LogMessage "Loading users" -VerboseOut
$usersIn = Import-Csv -Path $PATH_USERS_IN

$countTotalUsers = $usersIn.count

#Setting up users by operation action param
switch ($Operation) {
  "Enable" {
    $logMessage = "Setting personal sites for $($countTotalUsers) users into 'Read-Only' lockState"
    $lockState = 'ReadOnly'
  }
  "Disable" {
    $logMessage = "Setting personal sites for $($countTotalUsers) users into 'Unlock' lockState"
    $lockState = 'Unlock'
  }
}

Write-Log -Level INFO -LogMessage $logMessage -VerboseOut
$usersOut = Set-UserSites -UserList $usersIn -LockState $lockState


$totalSuccessfullyActions = ($usersOut | Where-Object { $_.PostConfigurationState -eq "ReadOnly" } | Measure-Object).Count

$usersOut | Format-Table UserPrincipalname, PostConfigurationState, PersonalSite

if ($totalSuccessfullyActions -gt 0) {
  
  $thisExportPath = "$($PATH_EXPORTS)\$($TODAY)_OneDriveReadOnly.csv"
  Write-Log -Level INFO -LogMessage "Export results in: $($thisExportPath)" -VerboseOut
  $usersOut | Export-Csv -Path $thisExportPath -NoTypeInformation

  #Put into globalrepository
  $usersOut | Export-Csv -Path "$($REPO_PATH)\$($TODAY)_$($Collection)_SetOneDriveReadonly.csv" -NoTypeInformation
}
else {
  Write-Log -Level INFO -LogMessage "No actions performed" -VerboseOut
}

