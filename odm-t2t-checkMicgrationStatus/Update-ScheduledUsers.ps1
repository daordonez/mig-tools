<#
  Date: 2023/03/03
  Author: Diego Ordonez
  Synopsis: Create local inventory of manual scheduled users on basis local migration team repository
#>



$PATH_SCRIPT = $PSScriptRoot
$PATH_CONFIG = "$($PATH_SCRIPT)\project_config.json"
$PATH_COMMON = "$($PSScriptRoot)\common"
$PATH_ODMEVENTS = "$($PATH_COMMON)\ODMGLobalEvents.csv"
$__CONFIG = Get-Content -Path $PATH_CONFIG | ConvertFrom-Json
$TODAY = Get-Date


#Invoke-Expression -Command "$($PATH_SCRIPT)\ConnectODMApi.ps1"  

#Create root directory Path
$pathPrefix = $__CONFIG.params.migrationFilesRepository.rootPathPrefix
$pathUser = (whoami).split('\')[1]
$pathSufix = $__CONFIG.params.migrationFilesRepository.rootPathSufix

#PUlling ODM Events
Write-Host "Updating ODM Events"
$GLOBALEV = Get-OdmEvent 
$GLOBALEV | Export-Csv -Path $PATH_ODMEVENTS -NoTypeInformation

$PATH_ROOT_REPOSITORY = "$($pathPrefix)\$($pathUser)\$($pathSufix)"
$MAIN_OUT = @()
#Test if repository local exists
if ( Test-Path -Path $PATH_ROOT_REPOSITORY ) {
  Write-Host "ROOT repository Found: $($PATH_ROOT_REPOSITORY)"
  Write-Host "Finding waves folders"
  #Get all current directories
  #Each directory represents a wave migration scheduled
  $ThisWaves = Get-ChildItem -Path $PATH_ROOT_REPOSITORY

  #Iterate folders
  $ThisWaves | ForEach-Object {
    $thisCollectionName = ""
    $thisCollectionUsers = @()
    $thisTempUsers = @()
    #Identify Wave
    Write-Host "# Identifiying user collection"
    $thisUsersCollection = Get-ChildItem -Path $_.FullName | where-object { $_.Name -like "*Collection.csv" }

    if ($null -ne $thisUsersCollection) {
      #Loading file
      $thisCollectionUsers = Import-Csv -Path $thisUsersCollection.FullName
      $thisCollectionName = ($thisCollectionUsers | Select-Object -Unique Collection).Collection

      Write-Host "# CollectionName: $($thisCollectionName)"
    }

    $thisUsersFile = Get-ChildItem -Path $_.FullName | where-object { $_.Name -like "*finalUsers.csv" }

    if ($null -ne $thisUsersFile) {
      Write-Host "# UserFile Found: $($thisUsersFile.Name)"
      Write-Host "> Importing user file: $($thisUsersFile.Name)" -ForegroundColor Yellow
      Import-Csv -Path $thisUsersFile.FullName -Delimiter ';' | ForEach-Object {
        #Create object for each user
        $thisNewUser = [PSCustomObject]@{
          LastRefreshTimeStamp      = $TODAY
          SourceUserPrincipalname   = $_.sourceUserPrincipalName
          TargetUserPrincipalName   = $_.targetUserPrincipalName
          MailboxType               = $_.mailBoxType
          Wave                      = $thisCollectionName
          WaveClosingTimeStamp      = $thisUsersFile.CreationTime
          isMailboxDisabled         = $false
          isOneDriveReadOnly        = ""
          OneDriveReadOnlyTimeStamp = ""
          DisableMailboxTimeStamp   = ""
          isInCollection            = $false
        }

        if ( $thisCollectionUsers.UserPrincipalName.contains($thisNewUser.SourceUserPrincipalname)) {
          $thisNewUser.isInCollection = $true
        }

        $thisTempUsers += $thisNewUser
      }

    }

    $thisUserBlockMailAccess = Get-ChildItem -Path $_.FullName | where-object { $_.Name -like "*DisableMailboxAccess.csv" }
    if ($null -ne $thisUserBlockMailAccess) {
      
      Write-Host "> Importing block file: $($thisUserBlockMailAccess.Name)" -ForegroundColor Yellow

      $BlockedUsers = Import-Csv -Path $thisUserBlockMailAccess.FullName

      $thisTempUsers | foreach-object {
        $thisUserToFind = $_
        $DisableStatus = $false
        $thisBlocked = $BlockedUsers | Where-Object { $_.UserPrincipalName -eq $thisUserToFind.SourceUserPrincipalname }

        switch ($thisBlocked.LicenseType) {
          "F3" {
            $DisableStatus = $thisBlocked.isWebClientsDisabled
          }
          "E3" {
            $DisableStatus = $thisBlocked.isAllClientsDisabled
          }
        }
        
        $_.isMailboxDisabled = $DisableStatus
        $_.DisableMailboxTimeStamp = $thisBlocked.TimeStamp

      } 
    }

    $thisUserReadOnly = Get-ChildItem -Path $_.FullName | where-object { $_.Name -like "*SetOneDriveReadonly.csv" }
    if ($null -ne $thisUserReadOnly) {
      Write-Host "> Importing ReadOnly file: $($thisUserReadOnly.Name)" -ForegroundColor Yellow
      $ODReadOnly = Import-Csv -Path $thisUserReadOnly.FullName

      $thisTempUsers | foreach-object {
        $thisODToFind = $_

        $thisReadONly = $ODReadOnly | Where-Object { $_.UserPrincipalName -eq $thisODToFind.SourceUserPrincipalname }

        $_.isOneDriveReadOnly = $thisReadONly.PostConfigurationState

        $_.OneDriveReadOnlyTimeStamp = $thisReadONly.ConfigurationTimeStamp

      }
    }
   


    $MAIN_OUT += $thisTempUsers

  }

  $MAIN_OUT | Format-Table -AutoSize

  $outPutPath = "$($PATH_COMMON)\LocalRepository_ScheduledUsers.csv"
  Write-Host "Exporting results. PATH:$($outPutPath)"
  $MAIN_OUT | Export-Csv -Path $outPutPath -NoTypeInformation

}
else {
  Write-Error "Fatal. 'root' directory NOT FOUND"
}
