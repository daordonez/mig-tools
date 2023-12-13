<#
  Date 23/01/2023
  Author: Diego Ordonez
  Synopsis: Disable user mailbox access (by Exchange protocols)
#>

param(
  [Parameter(Position = 0, Mandatory = $false)]
  [ValidateSet("All", "MobileDevices", "DesktopClients", "WebClients")]
  $AccessType = "All",
  [Parameter(Position = 1, Mandatory = $true)]
  [ValidateSet("Enable", "Disable")]
  $Action,
  [Parameter(Position = 2, Mandatory = $false)][switch]$RefreshToken,
  [Parameter(Position = 2, Mandatory = $true)]$Collection,
  [Parameter(Position = 3, Mandatory = $false)][switch]$CreateConnection
)


#Const
$PATH_SCRIPT = $PSScriptRoot
$PATH_CONFIG = "$($PATH_SCRIPT)\project_config.json"
$PATH_USERS_IN = "$($PSScriptRoot)\Users.csv"
$PATH_EXPORT = "$($PSScriptRoot)\exports"
$PATH_LOGS = "$($PATH_SCRIPT)\logs"
$__CONFIG = Get-Content -Path $PATH_CONFIG | ConvertFrom-Json
$TODAY = Get-Date
$LOGS_FILE = $__CONFIG.params.mailboxAccess.logsFileName
$PATH_EXEC_LOG = "$($PATH_LOGS)\$($LOGS_FILE)"

#Modules
Import-Module AzureAD
Import-Module ExchangeOnlineManagement

#Generate RepoPATH
$CurrentUser = (whoami).split("\\")[1]
$REPO_PATH = Join-Path -Path $__CONFIG.params.baseRepoPathPrefix -ChildPath "$($CurrentUser)\$($__CONFIG.params.baseReportPathSufix)"

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
    Write-Error "LogFile '$($LOGS_FILE)' not found"
  }
}


function Set-UserMailboxAccess {
  param(
    [Parameter(Position = 0, Mandatory = $true)]$UserPrincipalName,
    [Parameter(Position = 1, Mandatory = $true)]
    [ValidateSet("Enable", "Disable")]
    $ActionType,

    [Parameter(Position = 1, Mandatory = $true)]
    [ValidateSet("All", "MobileDevices", "DesktopClients", "WebClients")]
    $Type
  )

  $isProtocolEnabled = ""

  switch ($ActionType) {
    Enable {
      $isProtocolEnabled = $true
    }
    Disable {
      $isProtocolEnabled = $false
    }
  }

  switch ($Type) {
    All {
      Set-CasMailbox -Identity $UserPrincipalName `
        -EwsEnabled $isProtocolEnabled `
        -OwaEnabled $isProtocolEnabled `
        -OWAforDevicesEnabled $isProtocolEnabled `
        -MAPIEnabled $isProtocolEnabled `
        -EwsAllowOutlook $isProtocolEnabled `
        -OutlookMobileEnabled $isProtocolEnabled `
        -ActiveSyncEnabled $isProtocolEnabled
    }
    DesktopClients {
      Set-CasMailbox -Identity $UserPrincipalName `
        -MAPIEnabled $isProtocolEnabled `
        -EwsAllowOutlook $isProtocolEnabled `
      
    }
    MobileDevices {
      Set-CasMailbox -Identity $UserPrincipalName `
        -OutlookMobileEnabled $isProtocolEnabled `
        -ActiveSyncEnabled $isProtocolEnabled
    }
    WebClients {
      Set-CasMailbox -Identity $UserPrincipalName `
        -OWAforDevicesEnabled $isProtocolEnabled `
        -OwaEnabled $isProtocolEnabled
    }
  }
}


function Get-UserMailboxAccess {
  param(
    [Parameter(Position = 0, Mandatory = $true)]$UserPrincipalName,
    [Parameter(Position = 1, Mandatory = $true)]
    [ValidateSet("All", "MobileDevices", "DesktopClients", "WebClients")]
    $Type,
    [Parameter(Position = 2, Mandatory = $true)]
    [ValidateSet("Enable", "Disable")]
    $ActionType,
    [Parameter(Mandatory = $false)][switch]$OutputStatus
  )

  $userMailbox = Get-CasMailbox -Identity $UserPrincipalName

  $userTypeStatus = [PSCustomObject]@{
    TimeStamp                = $TODAY
    UserPrincipalName        = $UserPrincipalName
    Action                   = $ActionType
    LicenseType              = ""
    TokenRefresh             = $false
    RefreshTokenTimeStamp    = ""
    isAllClientsDisabled     = $false
    isDesktopClientsDisabled = $false
    isMobileDevicesDisabled  = $false
    isWebClientsDisabled     = $false
  }

  switch ($Type) {
    All {

      if ($OutputStatus -eq $true) {
        $userMailbox
      }
      
      #Desktop
      if (($userMailbox.MapiEnabled -eq $false) -and ($userMailbox.EwsAllowOutlook -eq $false)) {
        $userTypeStatus.isDesktopClientsDisabled = $true
      }
      #MobileDevice
      if (($userMailbox.OutlookMobileEnabled -eq $false) -and ($userMailbox.ActiveSyncEnabled -eq $false)) {
        $userTypeStatus.isMobileDevicesDisabled = $true
      }
      #WebClient
      if (($userMailbox.EwsEnabled -eq $false) -and ($userMailbox.OWAforDevicesEnabled -eq $false) -and ($userMailbox.OwaEnabled -eq $false)) {
        $userTypeStatus.isWebClientsDisabled = $true
      }

      #All
      if ($userTypeStatus.isDesktopClientsDisabled -and $userTypeStatus.isMobileDevicesDisabled -and $userTypeStatus.isWebClientsDisabled) {
        $userTypeStatus.isAllClientsDisabled = $true
      }
      
    }
    DesktopClients {
      if ($OutputStatus -eq $true) {
        $userMailbox | Format-Table MapiEnabled, EwsAllowOutlook
      }

      if (($userMailbox.MapiEnabled -eq $false) -and ($userMailbox.EwsAllowOutlook -eq $false)) {
        $userTypeStatus.isDesktopClientsDisabled = $true
      }
    }
    MobileDevices {
      if ($OutputStatus -eq $true) {
        $userMailbox | Format-Table OutlookMobileEnabled, ActiveSyncEnabled
      }

      if (($userMailbox.OutlookMobileEnabled -eq $false) -and ($userMailbox.ActiveSyncEnabled -eq $false)) {
        $userTypeStatus.isMobileDevicesDisabled = $true
      }
    }
    WebClients {
      if ($OutputStatus -eq $true) {
        $userMailbox | Format-Table EwsEnabled, OWAforDevicesEnabled, OwaEnabled
      }

      if (($userMailbox.EwsEnabled -eq $false) -and ($userMailbox.OWAforDevicesEnabled -eq $false) -and ($userMailbox.OwaEnabled -eq $false)) {
        $userTypeStatus.isWebClientsDisabled = $true
      }
    }
  }

  return $userTypeStatus
}

function Get-ExpectedResultOutput {
  param(
    [Parameter(Mandatory = $true)][ValidateSet("Enable", "Disable")]$ExpectedResult,
    [Parameter(Mandatory = $true)]$Property
  )

  $successString = ""
  $failString = ""
  $resultSelector
  #Result string set
  switch ($ExpectedResult) {
    Enable {
      $successString = "`tOK-Enabled"
      $failString = "`t! Disabled"
      $resultSelector = $false

        
    }
    Disable {
      $successString = "`tOK-Disabled"
      $failString = "`t! Enabled"
      $resultSelector = $true
    }
  }
  
  Start-Sleep -Seconds 1

  if ($Property -eq $resultSelector) {
    Write-Host $successString -ForegroundColor Green
  }
  else {
    Write-Host $failString -ForegroundColor Yellow
  }
}


Write-Log -Level INFO -LogMessage "Script start" -VerboseOut

#Connection Param
$isOKConnected = $false
$thisSourceDomain = $__CONFIG.params.sourceTenantDomain
if ($CreateConnection -eq $true) {

  Write-Log -Level INFO -LogMessage "Creating AAD connection. Tenant: $($thisSourceDomain)" -VerboseOut
  Connect-AzureAD
  Write-Log -Level INFO -LogMessage "Creating Exchange connection. Env: $($thisSourceDomain) (Source)" -VerboseOut
  Connect-ExchangeOnline
}
#Check Connection
try {
  
  $isAADConnected = (Get-AzureADTenantDetail).VerifiedDomains.name
  $isEXOConnected = (Get-ConnectionInformation).UserPrincipalName

  if ( ($isAADConnected.contains($thisSourceDomain)) -and ($isEXOConnected -like "*$($thisSourceDomain)")) {
    $isOKConnected = $true
  }

}
catch {
  Write-Log -Level ERROR -LogMessage "No needed connections"
  Write-Error "Please connect to ExhangeOnline and Azure AD, or use '-CreateConnection' param at script launching"
}


if ($isOKConnected) {

  #Write-Log -Level INFO -LogMessage "Project connfiguration loaded succesfully.SourceTenantDomain:$($__CONFIG.params.sourceTenantDomain),TargetTenantDomain:$($__CONFIG.params.targetTenantDomain)"
  Write-Log -Level INFO -LogMessage "Loading users" -VerboseOut
  $USERS_IN = Import-Csv -Path $PATH_USERS_IN
  
  $userCount = ($USERS_IN | Measure-Object).count
  
  Write-Log -Level INFO -LogMessage "User params: Action:$($Action),AccesType:$($AccessType),TotalUsers:$($userCount)" -VerboseOut
  
  Write-Log -Level INFO -LogMessage "User list:"
  
  #OutputUsers
  $USERS_IN | Select-Object SourceUserPrincipalName | Format-Table
  
  $userConfirmation = Read-Host "Proceed with mailbox access actions? [Y/n]"
  
  if (($userConfirmation -eq "y") -or ($userConfirmation -eq "Y")) {
  
    Write-Log -Level WARN -LogMessage "User confirmation:$($userConfirmation)"

    if ($RefreshToken -eq $true) {
      Write-Log -Level WARN -LogMessage "TokenRefresh Enabled." -VerboseOut
    }
    Write-Log -Level INFO -LogMessage "Deploying configuration..." -VerboseOut
    
    #start Execution
    $USERS_IN | ForEach-Object {
      $thisUPN = $_.SourceUserPrincipalName
      $thisLicenseType = $_.LicenseType
      switch ($AccessType) {
        All {
          #Action -> Enable/Disable
          #All -> AllMailboxProtocols

          #Analize license type
          switch ($thisLicenseType) {
            E3 {
              Set-UserMailboxAccess -UserPrincipalName $thisUPN -ActionType $Action -Type All
              #Write-Host "Disabling All for: UPN=$($thisUPN).LicenseType=$($thisLicenseType)"
            }
            F3 {
              Set-UserMailboxAccess -UserPrincipalName $thisUPN -ActionType $Action -Type WebClients
              #Write-Host "Disabling webclient for: UPN=$($thisUPN).LicenseType=$($thisLicenseType)"
            }
          }

        }
        DesktopClients {
          Set-UserMailboxAccess -UserPrincipalName $thisUPN -ActionType $Action -Type DesktopClients
        }
        MobileDevices {
          Set-UserMailboxAccess -UserPrincipalName $thisUPN -ActionType $Action -Type MobileDevices
        }
        WebClients {
          Set-UserMailboxAccess -UserPrincipalName $thisUPN -ActionType $Action -Type WebClients
        }
      }
    }
  
    Write-Log -Level INFO -LogMessage "Gathering deployment Status" -VerboseOut
    $allUserResults = @()
  
    #Collect deployment status
    $USERS_IN | ForEach-Object {
      Write-Host "User: $($_.SourceUserPrincipalName)"
      $LicenseTypeCheck = $_.LicenseType
      #Return custom object
      #Properties: 
  
      $userResults = Get-UserMailboxAccess -UserPrincipalName $_.SourceUserPrincipalName -Type All -ActionType $Action
  
      switch ($AccessType) {
        All {
          
          switch ($LicenseTypeCheck) {
            E3 {

              Write-Host "`t - DesktopClients:" -NoNewline
              Get-ExpectedResultOutput -ExpectedResult $Action -Property $userResults.isDesktopClientsDisabled
      
              Write-Host "`t - MobileDevices:" -NoNewline
              Get-ExpectedResultOutput -ExpectedResult $Action -Property $userResults.isMobileDevicesDisabled
      
              Write-Host "`t - WebClients:`t" -NoNewline
              Get-ExpectedResultOutput -ExpectedResult $Action -Property $userResults.isWebClientsDisabled
            }
            F3 {
              Write-Host "`t - WebClients:`t" -NoNewline
              Get-ExpectedResultOutput -ExpectedResult $Action -Property $userResults.isWebClientsDisabled
            }
          }

  
        }
        DesktopClients {
        
          Write-Host "`t - DesktopClients:" -NoNewline
          Get-ExpectedResultOutput -ExpectedResult $Action -Property $userResults.isDesktopClientsDisabled
  
        }
        MobileDevices {
        
          Write-Host "`t - MobileDevices:" -NoNewline
          Get-ExpectedResultOutput -ExpectedResult $Action -Property $userResults.isMobileDevicesDisabled
  
        }
        WebClients {
  
          Write-Host "`t - WebClients:`t" -NoNewline
          Get-ExpectedResultOutput -ExpectedResult $Action -Property $userResults.isWebClientsDisabled
  
        }
      }
      
      #Refresh token. It will force user sign in
      if ($RefreshToken -eq $true) {
  
        Write-Host "`t - TokenRefresh:`t$true" -ForegroundColor Yellow
        Revoke-AzureADUserAllRefreshToken -ObjectId $thisUPN
        #Get Last 'RefreshTokensValidFromDateTime' and store
        $userResults.TokenRefresh = (Get-AzureADUser -ObjectId $thisUPN).RefreshTokensValidFromDateTime
  
        $userResults.TokenRefresh = $true
      }
      
      #Save license Detail
      $userResults.LicenseType = $LicenseTypeCheck
      $allUserResults += $userResults
    }
  
    Write-Log -Level INFO -LogMessage "User Summary:" -VerboseOut
    $allUserResults | Format-List *
    $todayToStr = $TODAY | Get-Date -Format "yyyyMMddHHmm"
    Write-Log -Level INFO -LogMessage "Exporting results into 'exports' directory" -VerboseOut
    $allUserResults | Export-Csv -Path "$($PATH_EXPORT)\$($todayToStr)_MailboxAccess.csv" -NoTypeInformation

    #Put into globalrepository
    $allUserResults | Export-Csv -Path "$($REPO_PATH)\$($todayToStr)_$($Collection)_DisableMailboxAccess.csv" -NoTypeInformation
  }
  else {
    Write-Log -Level INFO -LogMessage "Exit. Script End"
  }
}





