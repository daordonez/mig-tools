<#
    Date: 2022/12/22
    Environment: Testing
    Author: Avanade migration Team
    Synopsis: This script will copy all regional settings configurations of source user in target user  ( different O365 Tenants)
#>

param(
  # Skip user interaction in copy scenario
  [Parameter(Mandatory = $false)]
  [switch]
  $Confirmation,
  # Read-Only
  [Parameter(Mandatory = $false)]
  [switch]
  $DisplayCurrentStatus
)

$TODAY = Get-Date
$PATH_SCRIPT = $PSScriptRoot
$env:PATH_LOGS = "$($PATH_SCRIPT)\logs"
$env:PATH_BACKUP = "$($PATH_SCRIPT)\backups"
$env:PATH_CONFIG = "$($PATH_SCRIPT)\project_config.json"
$env:PATH_EXEC_LOG = "$($env:PATH_LOGS)\CopyRegionalSettings.log"
$PATH_USERS_IN = "$($PATH_SCRIPT)\Users.csv"

#Logging
function Write-Log {

  param(
    [Parameter(Mandatory = $true)][ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]$Level,
    [Parameter(Mandatory = $true)]$LogMessage,
    [Parameter(Mandatory = $false)][switch]$VerboseOut
      
  )

  $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")

  $Line = "$Stamp $Level $LogMessage"

  if (Test-Path -Path $env:PATH_EXEC_LOG) {

    Add-Content -Path $env:PATH_EXEC_LOG -Value $Line

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

    Write-Error "LogFile 'CopyRegionalSettings.log' not found"
  }

}

function New-UserMailboxes {
  
  param(
    [Parameter(Mandatory = $true)]$SourceUserPrincipalName,
    [Parameter(Mandatory = $true)]$TargetUserPrincipalName,
    [Parameter(Mandatory = $true)]$SourceConfiguration,
    [Parameter(Mandatory = $true)]$TargetConfiguration
  )

  $thisMailboxRegionalConf = [PSCustomObject]@{
    SourceUserPrincipalName                     = $SourceUserPrincipalName
    TargetUserPrincipalName                     = $TargetUserPrincipalName
    SourceDateFormat                            = $sourceConfiguration.DateFormat
    SourceLanguage                              = $sourceConfiguration.Language
    SourceDefaultFolderNameMatchingUserLanguage = $sourceConfiguration.DefaultFolderNameMatchingUserLanguage
    SourceTimeFormat                            = $sourceConfiguration.TimeFormat
    SourceTimeZone                              = $sourceConfiguration.TimeZone
    SourceIdentity                              = $sourceConfiguration.Identity
    SourceIsValid                               = $sourceConfiguration.IsValid
    SourceObjectState                           = $sourceConfiguration.ObjectState
    TargetDateFormat                            = $targetConfiguration.DateFormat
    TargetLanguage                              = $targetConfiguration.Language
    TargetDefaultFolderNameMatchingUserLanguage = $targetConfiguration.DefaultFolderNameMatchingUserLanguage
    TargetTimeFormat                            = $targetConfiguration.TimeFormat
    TargetTimeZone                              = $targetConfiguration.TimeZone
    TargetIdentity                              = $targetConfiguration.Identity
    TargetIsValid                               = $targetConfiguration.IsValid
    TargetObjectState                           = $targetConfiguration.ObjectState
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
    [Parameter(Mandatory = $true)]$SourceUserMailbox,
    [Parameter(Mandatory = $true)]$TargetUsermailbox
  )

  $isCopied = $false

  $ref = $SourceUserMailbox[1]
  $dif = $TargetUsermailbox[1]

  $comparisionResult = Compare-Object -ReferenceObject $ref -DifferenceObject $dif -IncludeEqual -Property Language, TimeZone

  if ( ($comparisionResult.SideIndicator.contains("==")) -and ($comparisionResult.Count -eq 1) ) {
    $isCopied = $true
  }

  return $isCopied
}

function Get-CurrentConfiguration {

  param(
    [Parameter(Mandatory = $true)]$SourceUPN,
    [Parameter(Mandatory = $true)]$TargetUPN
  )

  #Gathering source configuration
  $sourceConfigurationAll = Get-RegionalConfig -UserPrincipalname $SourceUPN -Environment SRC
 
  #Gathering target configuration
  $targetConfigurationAll = Get-RegionalConfig -UserPrincipalname  $TargetUPN -Environment TRG

  #Object
  $thisMailboxRegionalConf = New-UserMailboxes -SourceConfiguration $sourceConfigurationAll -TargetConfiguration $targetConfigurationAll -SourceUserPrincipalName $SourceUPN -TargetUserPrincipalName $TargetUPN

  #Checking current status
  $thisMailboxRegionalConf.ConfigurationCopied = Get-ConfigurationCopyStatus -SourceUserMailbox $sourceConfigurationAll -TargetUsermailbox $targetConfigurationAll

  return $thisMailboxRegionalConf
}

function Copy-RegionalConfiguration {

  param(
    [Parameter(Mandatory = $true)]$UserArray
  )

  Write-Log -Level INFO -LogMessage "Launching mailbox regional configuration copy" -VerboseOut
  
  #Dumping Array to tmp
  $thisWorkArray = @()

  $UserArray | ForEach-Object {
  
    #Setting mailbox with source params
    if(($null -ne $_.SourceLanguage) -or ($null -ne $_.SourceDateFormat)){
      Set-TRGMailboxRegionalConfiguration -Identity $_.TargetUserPrincipalName -DateFormat $_.SourceDateFormat -Language $_.SourceLanguage -TimeFormat $_.SourceTimeFormat -TimeZone $_.SourceTimeZone -LocalizeDefaultFolderName:$true
    }
  }
  
  #Waiting for propagation
  Write-Log -Level INFO -LogMessage "Mailbox configuration copy done. Waiting propagation" -VerboseOut
  Start-Sleep -Seconds 30
  
  #Checking CopyStatus
  $UserArray | ForEach-Object {
  
    $thisSourceUserConfig = Get-RegionalConfig -UserPrincipalname $_.SourceUserPrincipalName -Environment SRC
    $thisTargetUserConfig = Get-RegionalConfig -UserPrincipalname $_.TargetUserPrincipalName -Environment TRG

    #Object

    $tmpObject = New-UserMailboxes -SourceConfiguration $thisSourceUserConfig -TargetConfiguration $thisTargetUserConfig -SourceUserPrincipalName $_.SourceUserPrincipalName -TargetUserPrincipalName $_.TargetUserPrincipalName
  
    $tmpObject.ConfigurationCopied = Get-ConfigurationCopyStatus -SourceUserMailbox $thisSourceUserConfig -TargetUsermailbox $thisTargetUserConfig

    $thisWorkArray += $tmpObject
  }

  return $thisWorkArray

}

function New-ConnectionSet {
  Write-Log -Level INFO -LogMessage "Creating SOURCE ($($__CONFIG.params.sourceTenantDomain)) connection" -VerboseOut
  Connect-ExchangeOnline -Prefix "SRC" -ShowBanner:$false

  Write-Log -Level INFO -LogMessage "Creating TARGET ($($__CONFIG.params.targetTenantDomain)) connection" -VerboseOut
  Connect-ExchangeOnline -Prefix "TRG" -ShowBanner:$false
}

function Get-ConnectionStatus {

  $isConnectionsSuccess = $false
  $EXOconnectionStatus = Get-ConnectionInformation

  $sourceConnection = $false
  $targetConnection = $false

  if ($null -ne $EXOconnectionStatus) {

    $EXOconnectionStatus | ForEach-Object {
      if (($_.UserPrincipalname -like "*@$($__CONFIG.params.sourceTenantDomain)") -and ($_.ModulePrefix -like "$($__CONFIG.params.sourceConnectionPrefix)") ) {
        $sourceConnection = $true
      }
  
      if (($_.UserPrincipalname -like "*@$($__CONFIG.params.targetTenantDomain)") -and ($_.ModulePrefix -like "$($__CONFIG.params.targetConnectionPrefix)")) {
        $targetConnection = $true
      }
    }
  }


  if ($sourceConnection -and $targetConnection -and ($null -ne $EXOconnectionStatus)) {
    $isConnectionsSuccess = $true
  }

  return $isConnectionsSuccess
}



Write-Log -Level INFO -LogMessage "Script start" -VerboseOut
$__CONFIG = Get-Content -Path $env:PATH_CONFIG | ConvertFrom-Json

if ($null -ne $__CONFIG) {
  Write-Log -Level INFO -LogMessage "Project connfiguration loaded succesfully.SourceTenantDomain:$($__CONFIG.params.sourceTenantDomain),TargetTenantDomain:$($__CONFIG.params.targetTenantDomain)"

  #Collecting environments connections
  $ConnectionsStatus = Get-ConnectionStatus

  if ($ConnectionsStatus -eq $true) {
    Write-Log -Level INFO -LogMessage "Connections OK (Source & target)" -VerboseOut

    $EXOconnectionStatus | Format-List UserPrincipalname, ConnectionUri, TokenExpiryTimeUTC
  }
  else {
    Write-Log -Level WARN -LogMessage "Disconnecting old EXO sessions" -VerboseOut
    Disconnect-ExchangeOnline -Confirm:$false

    New-ConnectionSet

    $ConnectionsStatus = Get-ConnectionStatus
  }
  
  if ($ConnectionsStatus -eq $true) {

    $USERS = Import-Csv -Path $PATH_USERS_IN
 
    $usersCount = ($USERS | Measure-Object).Count
 
    Write-Log -Level INFO -LogMessage "Total users count in this batch: $($usersCount)"

    Write-Log -Level INFO -LogMessage "Gathering mailbox regional configuration. Both environments" -VerboseOut

    $allMbxRegionalConfigurationData = @()

    $USERS | ForEach-Object {
 
      $thisMailboxRegionalConf = Get-CurrentConfiguration -SourceUPN $_.SourceUserPrincipalName -TargetUPN $_.TargetUserPrincipalName

      $allMbxRegionalConfigurationData += $thisMailboxRegionalConf
      
    }

    #output users
    $usersToCopy = @()
    if ($DisplayCurrentStatus -eq $true) {
      $allMbxRegionalConfigurationData | Format-Table TargetUserPrincipalName, SourceLanguage, TargetLanguage, TargetTimeZone, ConfigurationCopied
    }
    else {
      $usersToCopy = $allMbxRegionalConfigurationData | Where-Object { $_.ConfigurationCopied -eq $false }
      $usersToCopy | Format-Table TargetUserPrincipalName, SourceLanguage, TargetLanguage, TargetTimeZone, ConfigurationCopied

      #Create Backup
      $bkpDate = $TODAY | Get-Date -Format "yyyyMMddHHmm"
      $allMbxRegionalConfigurationData | Export-Csv -Path "$($env:PATH_BACKUP)\bkp_$($bkpDate)_MailboxRegionalSettings.csv" -NoTypeInformation 
      Write-Log -Level INFO -LogMessage "Total users to copy: $($usersToCopy.count)"

      if ($usersToCopy.count -eq 0 ) {
        Write-Log -Level INFO -LogMessage "No users found to copy" -VerboseOut
      }
      else {
        if ($Confirmation -eq $true) {
        
          $tmpCopyStatus = Copy-RegionalConfiguration -UserArray $usersToCopy

          $usersToCopy = $tmpCopyStatus
  
          $totalNotCopied = ($tmpCopyStatus | Where-Object { $_.ConfigurationCopied -eq $false } | Measure-Object).Count
          $totalSuccess = ($tmpCopyStatus | Where-Object { $_.ConfigurationCopied -eq $true } | Measure-Object).Count
  
          if ($totalNotCopied -eq 0) {
            Write-Log -Level INFO -LogMessage "All users have been copied. Summary Copy:$($totalSuccess)/$($usersToCopy.count)" -VerboseOut

            $tmpCopyStatus | Format-Table TargetUserPrincipalName, SourceLanguage, TargetLanguage, TargetTimeZone, ConfigurationCopied
          }
  
        }
        else {
        
          $userConfirmation = Read-Host "Proceed with regional configuration Copy? [Y/n]"

          if (($userConfirmation -eq "y") -or ($userConfirmation -eq "Y")) {

            Write-Log -Level INFO -LogMessage "User confirmation:$($userConfirmation)"
            $tmpCopyStatus = Copy-RegionalConfiguration -UserArray $usersToCopy
  
            $totalNotCopied = ($tmpCopyStatus | Where-Object { $_.ConfigurationCopied -eq $false } | Measure-Object).Count
            $totalSuccess = ($tmpCopyStatus | Where-Object { $_.ConfigurationCopied -eq $true } | Measure-Object).Count
  
            if ($totalNotCopied -eq 0) {
              Write-Log -Level INFO -LogMessage "All users have been copied. Summary Copy:$($totalSuccess)/$($usersToCopy.count)" -VerboseOut

              $tmpCopyStatus | Format-Table TargetUserPrincipalName, SourceLanguage, TargetLanguage, TargetTimeZone, ConfigurationCopied
            }
          }
          else {
            Write-Log -Level INFO -LogMessage "User confirmation:$($userConfirmation). Cancelled"
          }
        }
      }
    }
  }
  else {
    Write-Log -Level WARN -LogMessage "One of the needed connections is not created yet. Please connect to source and target Exchange Online Service" -VerboseOut
  }
}
else {
  Write-Log -Level FATAL -LogMessage "Project configuration file missed or not found in script root directory" -VerboseOut
}

Write-Log -Level INFO -LogMessage "Script end" -VerboseOut





