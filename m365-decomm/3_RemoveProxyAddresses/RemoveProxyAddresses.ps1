param(
  [Parameter(Mandatory=$true)]$ToRemoveDomain,
  [Parameter(Mandatory=$true)]$ReplacementDomain
)

#Const
$PATH_ROOT = $PSScriptRoot
$TODAY = Get-Date -Format "yyyyMMddHHmm"
$PATH_EXEC_LOG = "$($PATH_ROOT)\log\RemoveProxyAddresses.log"
$PATH_OUTPUT = "$($PATH_ROOT)\output"
$PATH_USERSIN = "$($PATH_ROOT)\Users.csv"



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

    Write-Error "LogFile 'RemoveProxyAddresses.log' not found"
  }

}


#Runtime
Write-Log -Level INFO -LogMessage "Script RemoveProxyAddresses start" -VerboseOut
Write-Log -Level INFO -LogMessage "Input Params: ToRemoveDomain=$($ToRemoveDomain). ReplacementDomain=$($ReplacementDomain)" -VerboseOut
Write-Log -Level INFO -LogMessage "Loading users from CSV" -VerboseOut
$UsersIn = Import-Csv -Path $PATH_USERSIN
$UsersCount = ($UsersIn | Measure-Object).Count
Write-Log -Level INFO -LogMessage "Users to modify.TotalCount=$($UsersCount)" -VerboseOut


Write-Log -Level INFO -LogMessage "Starting removal of proxyAddresses" -VerboseOut
Write-Log -Level INFO -LogMessage "Removing 'proxyAddress' that contains '$($ToRemoveDomain)'" -VerboseOut

$UsersIn | ForEach-Object {

  $ThisPrimarySmtpAddress = $_.PrimarySmtpAddress

  switch ($_.RecipientType) {
    UserMailbox{
      $AddressesInMailbox = Get-Mailbox -Identity $ThisPrimarySmtpAddress -ErrorAction SilentlyContinue

      if($null -ne $AddressesInMailbox){
        $addressesRemain = $AddressesInMailbox.EmailAddresses | Where-Object {$_ -notlike "*@$($ToRemoveDomain)"}

        Set-Mailbox -Identity $ThisPrimarySmtpAddress -EmailAddresses $addressesRemain
      }
    }
    MailUniversalDistributionGroup{
      $addressesDL = Get-DistributionGroup -Identity $ThisPrimarySmtpAddress

      if($null -ne $addressesDL){
        $addressesRemainDL = $addressesDL.EmailAddresses | Where-Object {$_ -notlike "*@$($ToRemoveDomain)"}

        Set-DistributionGroup -Identity $ThisPrimarySmtpAddress -EmailAddresses $addressesRemainDL
      }
    }
    MailUniversalSecurityGroup{
      $addressesDG = Get-DistributionGroup -Identity $ThisPrimarySmtpAddress

      if($null -ne $addressesDL){
        $addressesRemainDG = $addressesDG.EmailAddresses | Where-Object {$_ -notlike "*@$($ToRemoveDomain)"}

        Set-DistributionGroup -Identity $ThisPrimarySmtpAddress -EmailAddresses $addressesRemainDG
      }
    }
    MailUser{
      $addressesMU = Get-MailUser -Identity $ThisPrimarySmtpAddress

      if($null -ne $addressesMU){
        $addressesRemainMU = $addressesMU.EmailAddresses | Where-Object {$_ -notlike "*@$($ToRemoveDomain)"}
        Set-MailUser -Identity $ThisPrimarySmtpAddress -EmailAddresses $addressesRemainMU
      }

    }
  }
}