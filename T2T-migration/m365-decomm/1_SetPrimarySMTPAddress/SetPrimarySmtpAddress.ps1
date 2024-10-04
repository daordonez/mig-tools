param(
  [Parameter(Mandatory=$true)]$ToRemoveDomain,
  [Parameter(Mandatory=$true)]$ReplacementDomain,
  [Parameter(Mandatory=$false)][switch]$SkipBackup
)

#Const
$PATH_ROOT = $PSScriptRoot
$TODAY = Get-Date -Format "yyyyMMddHHmm"
$PATH_EXEC_LOG = "$($PATH_ROOT)\log\SetPrimarySmtpAddress.log"
$PATH_BACKUP = "$($PATH_ROOT)\backup"
$PATH_OUTPUT = "$($PATH_ROOT)\output"

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

    Write-Error "LogFile 'SetPrimarySmtpAddress.log' not found"
  }

}

#Functions
function Set-MailboxDeployment{
  param(
    [Parameter(Mandatory=$true)]$PrimarySmtpAddress
  )
  
  $ThisMailboxAddress = Get-Mailbox -Identity $PrimarySmtpAddress -ErrorAction SilentlyContinue
  
    if($null -ne $ThisMailboxAddress){

      #Get 'onmicrosoft' address
      $thisNewAddress = ""

      $thisAddresses = $ThisMailboxAddress.EmailAddresses | Where-Object {$_ -like "*@$($ReplacementDomain)"}

      if(($thisAddresses | Measure-Object).Count -gt 1){
        $thisNewAddress = $thisAddresses[0]
      }else{
        $thisNewAddress = $ThisMailboxAddress.EmailAddresses | Where-Object {$_ -like "*@$($ReplacementDomain)"}
      }

      #Clear smtp prefix
      $newAddressClearPrefix = $thisNewAddress.split(':')[1]


      if($null -ne $newAddressClearPrefix){
        
        $list = @()
        $ThisMailboxAddress.EmailAddresses | ForEach-Object{
          #split parts
          $addressPrefix = $_.split(':')[0]
          $addressRemain = $_.split(':')[1]

          if($addressRemain.ToLower() -eq $newAddressClearPrefix.ToLower()){
            $address = "SMTP:" + $addressRemain
          }else{
            $address = $addressPrefix.ToLower() + ":" + $addressRemain
          }
          $list += $address
        }

        #Proceed with set
        Set-Mailbox -Identity $PrimarySmtpAddress -EmailAddresses $list

      }
    }

}

function Set-DistributionGroupDeployment{
  param(
    [Parameter(Mandatory=$true)]$PrimarySmtpAddress
  )

  $thisDistribution = Get-DistributionGroup -Identity $PrimarySmtpAddress

  if($null -ne $thisDistribution){
    $addressPrefix = $thisDistribution.PrimarySmtpAddress.split('@')[0]
    $newPrimary = "$($addressPrefix)@$($ReplacementDomain)"
    Set-DistributionGroup -Identity $PrimarySmtpAddress -PrimarySmtpAddress $newPrimary
  }

}

function Set-MailUserDeployment {
  param(
    [Parameter(Mandatory = $true)]$PrimarySmtpAddress
  )

  $thisMailUser = Get-MailUser -Identity $PrimarySmtpAddress

  if ($null -ne $thisMailUser) {
    #Get 'onmicrosoft' address
    $thisNewAddress = ""

    $thisAddresses = $thisMailUser.EmailAddresses | Where-Object { $_ -like "*@$($ReplacementDomain)" }

    if (($thisAddresses | Measure-Object).Count -gt 1) {
      $thisNewAddress = $thisAddresses[0]
    }
    else {
      $thisNewAddress = $thisMailUser.EmailAddresses | Where-Object { $_ -like "*@$($ReplacementDomain)" }
    }

    #Clear smtp prefix
    $newAddressClearPrefix = $thisNewAddress.split(':')[1]


    if ($null -ne $newAddressClearPrefix) {
      
      $listMailUser = @()
      $thisMailUser.EmailAddresses | ForEach-Object {
        #split parts
        $addressPrefix = $_.split(':')[0]
        $addressRemain = $_.split(':')[1]

        if ($addressRemain.ToLower() -eq $newAddressClearPrefix.ToLower()) {
          $address = "SMTP:" + $addressRemain
        }
        else {
          $address = $addressPrefix.ToLower() + ":" + $addressRemain
        }
        $listMailUser += $address
      }

      #Proceed with set
      Set-MailUser -Identity $PrimarySmtpAddress -EmailAddresses $listMailUser
    }
  }
}


#Runtime
Write-Log -Level INFO -LogMessage "Script SetPrimarySmtpAddress start" -VerboseOut
Write-Log -Level INFO -LogMessage "Input Params: ToRemoveDomain=$($ToRemoveDomain). ReplacementDomain=$($ReplacementDomain)" -VerboseOut
Write-Log -Level INFO -LogMessage "Loading users from CSV" -VerboseOut
$UsersIn = Import-Csv -Path .\Users.csv
$UsersCount = ($UsersIn | Measure-Object).Count
Write-Log -Level INFO -LogMessage "Users to modify.TotalCount=$($UsersCount)" -VerboseOut

#Backup
if($SkipBackup -ne $true){

  Write-Log -Level INFO -LogMessage "Creating Proxyaddress Backup" -VerboseOut
  
  $AddressToBackup = @()
  $UsersIn | ForEach-Object {
    #Object
    $ThisAddressBackup = [PSCustomObject]@{
      PrimarySmtpAddress = $_.PrimarySmtpAddress
      RecipientType = ""
      EmailAddresses = ""
    }
  
    $ThisRecipient = Get-Recipient -Identity $($_.PrimarySmtpAddress) -ErrorAction SilentlyContinue
  
    if($null -ne $ThisRecipient){
      $ThisAddressBackup.RecipientType = $ThisRecipient.RecipientType
      $ThisAddressBackup.EmailAddresses = $ThisRecipient.EmailAddresses -join ';'
    }else{
      $ThisAddressBackup.EmailAddresses = "no_mailbox_found"
  
    }
  
  
    $AddressToBackup += $ThisAddressBackup
  }
  
  $parseToRemoveDomain = $ToRemoveDomain -replace '\.','_'
  $OutputPath = "$($PATH_BACKUP)\$($TODAY)_AddressBackup_$($parseToRemoveDomain).csv"
  Write-Log -Level INFO -LogMessage "Exporting backup.PATH=$($OutputPath)" -VerboseOut
  
  $AddressToBackup | Export-Csv -Path $OutputPath -NoTypeInformation
}else{
  Write-Log -Level INFO -LogMessage "'SkipBackup' enabled. Backup will be skipped"
}

#Config deployment

Write-Log -Level INFO -LogMessage "Starting recipient(s) configuration" -VerboseOut
Write-Log -Level INFO -LogMessage "Setting 'primarySmtpAddress' to '$($ReplacementDomain)'" -VerboseOut

$UsersIn | ForEach-Object {

  $ThisPrimarySmtpAddress = $_.PrimarySmtpAddress

  switch ($_.RecipientType) {
    UserMailbox{
      Set-MailboxDeployment -PrimarySmtpAddress $ThisPrimarySmtpAddress
    }
    MailUniversalDistributionGroup{
      Set-DistributionGroupDeployment -PrimarySmtpAddress $ThisPrimarySmtpAddress
    }
    MailUniversalSecurityGroup{
      Set-DistributionGroupDeployment -PrimarySmtpAddress $ThisPrimarySmtpAddress
    }
    MailUser{
      Set-MailUserDeployment -PrimarySmtpAddress $ThisPrimarySmtpAddress
    }
  }
}
Write-Log -Level INFO -LogMessage "Recipient(s) deployment completed" -VerboseOut

#Config validation
$DeploymentObjects = @()
Write-Log -Level INFO -LogMessage "Getting current deployment status" -VerboseOut
Write-Log -Level INFO -LogMessage "Getting 'PrimarySmtpAddress'" -VerboseOut

$UsersIn | ForEach-Object {
  $ThisPrimarySmtpAddressCheck = $_.PrimarySmtpAddress

  $thisChecked = [PSCustomObject]@{
    OldPrimarySmtpAddress = $_.PrimarySmtpAddress
    CurrentPrimarySmtpAddress = ""
    RecipientType = ""
    Completed = $false
  }

  switch ($_.RecipientType) {
    UserMailbox{
      $ThisCheckMailbox = Get-Mailbox -Identity $ThisPrimarySmtpAddressCheck -ErrorAction SilentlyContinue

      if ($null -ne $ThisCheckMailbox) {
        $thisChecked.CurrentPrimarySmtpAddress = $ThisCheckMailbox.PrimarySmtpAddress
        $thisChecked.RecipientType = $_.RecipientType
        if($ThisCheckMailbox.PrimarySmtpAddress -like "*$($ReplacementDomain)"){
          $thisChecked.Completed = $true
        }
      }
    }
    MailUniversalDistributionGroup{
      $ThisCheckDG = Get-DistributionGroup -Identity $ThisPrimarySmtpAddressCheck -ErrorAction SilentlyContinue

      if ($null -ne $ThisCheckDG) {
        $thisChecked.CurrentPrimarySmtpAddress = $ThisCheckDG.PrimarySmtpAddress
        $thisChecked.RecipientType = $_.RecipientType
        if($ThisCheckDG.PrimarySmtpAddress -like "*$($ReplacementDomain)"){
          $thisChecked.Completed = $true
        }
      }
    }
    MailUniversalSecurityGroup{
      $ThisCheckDG = Get-DistributionGroup -Identity $ThisPrimarySmtpAddressCheck -ErrorAction SilentlyContinue

      if ($null -ne $ThisCheckDG) {
        $thisChecked.CurrentPrimarySmtpAddress = $ThisCheckDG.PrimarySmtpAddress
        $thisChecked.RecipientType = $_.RecipientType
        if($ThisCheckDG.PrimarySmtpAddress -like "*$($ReplacementDomain)"){
          $thisChecked.Completed = $true
        }
      }
    }
    MailUser{
      $ThisCheckMU = Get-MailUser -Identity $ThisPrimarySmtpAddressCheck -ErrorAction SilentlyContinue

      if ($null -ne $ThisCheckMU) {
        $thisChecked.CurrentPrimarySmtpAddress = $ThisCheckMU.PrimarySmtpAddress
        $thisChecked.RecipientType = $_.RecipientType
        if($ThisCheckMU.PrimarySmtpAddress -like "*$($ReplacementDomain)"){
          $thisChecked.Completed = $true
        }
      }
    }
  }

  $DeploymentObjects += $thisChecked
}

Write-Log -Level INFO -LogMessage "Output results in UI" -VerboseOut
$parseToRemoveDomain = $ToRemoveDomain -replace '\.','_'
$OutputResults = "$($PATH_OUTPUT)\$($TODAY)_SetPrimarySmtpAddress_$($parseToRemoveDomain).csv"
Write-Log -Level INFO -LogMessage "Output results exported. PATH=$($OutputResults)" -VerboseOut
Write-Log -Level INFO -LogMessage "Script end" -VerboseOut

$DeploymentObjects | Out-GridView
$DeploymentObjects | Export-Csv -Path $OutputResults -NoTypeInformation




