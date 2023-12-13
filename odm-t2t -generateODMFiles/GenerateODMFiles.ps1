
param(
  [Parameter(Position = 0, Mandatory = $true)]$CollectionName,
  [Parameter(Position = 1, Mandatory = $false)][switch]$GenerateFromIAMFile
)

$ROOT_SCRIPT = $PSScriptRoot
$TODAY = Get-Date -Format "yyyyMMdd"
$PATH_FILES = "$($ROOT_SCRIPT)\files"
$thisWorkingDir
$IMPORTABLE_CSV_PATH = ""
$CurrentUser = (whoami).split("\\")[1]
$IAMRepositoryPath = "C:\Users\$($CurrentUser)\OneDrive - Avanade\Unishare Slovakia\Migration\Waves"

if ($GenerateFromIAMFile -eq $true) {
  #Set IAM source Path
  Write-Warning "Generating from IAM repository"
  $CurrentUser = (whoami).split("\\")[1]
  $IAMRepositoryPath = "C:\Users\$($CurrentUser)\OneDrive - Avanade\Unishare Slovakia\Migration\Waves"

  
  if (Test-Path -Path $IAMRepositoryPath) {
    
    Write-Host "IAMRepository: $($IAMRepositoryPath)"
    $WorkingDir = Get-ChildItem -Path $IAMRepositoryPath | Where-Object { $_.Name -like "*$($CollectionName)*" } | Sort-Object LastWriteTime

    
    #Check if for a singleton directory. In case more than one, choose the one with the last write time
    $totalWorkingDirs = ($WorkingDir | Measure-Object).Count
    if ($totalWorkingDirs -gt 1) {

      #Check if directory existis
      if (Test-Path -Path $($WorkingDir[1].FullName)) {
        $thisWorkingDir = $WorkingDir[1]

      }
      
    }
    else {
      $thisWorkingDir = $WorkingDir
    }
    
    Write-Host "WorkinkDirectory : $($thisWorkingDir.Name)"

    #Check finalUsers file
    $thisFinalUsersFile = Get-ChildItem -Path $thisWorkingDir.FullName | Where-Object { $_.Name -like "*$($CollectionName)_finalUsers*" }

    if (($null -ne $thisFinalUsersFile) -and (Test-Path -Path $thisFinalUsersFile.FullName)) {
      Write-Host "`t IAM finalUsersFile Found: $($thisFinalUsersFile.Name)" -ForegroundColor Green
      
      $IMPORTABLE_CSV_PATH = $thisFinalUsersFile.FullName
    }


  }
}
else {
  $IMPORTABLE_CSV_PATH = "$($ROOT_SCRIPT)\Users.csv"
}

Write-Host "loading users from file"
$UsersIn = Import-Csv -Path $IMPORTABLE_CSV_PATH -Delimiter ';'

$DiscoverFile = @()
$MappingFile = @()
$CollectionFile = @()

$UsersIn | ForEach-Object {

  Write-Host "Generating user: $($_.SourceUserPrincipalName)"

  #DiscoveryFile Object
  $thisDiscoveryObject = [PSCustomObject]@{
    UserPrincipalname = $_.SourceUserPrincipalName
    Type              = "User"
  }

  $DiscoverFile += $thisDiscoveryObject

  #MappingFile Object
  $thisMappingObject = New-Object -TypeName PSObject
  $thisMappingObject | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $_.SourceUserPrincipalName
  $thisMappingObject | Add-Member -MemberType NoteProperty -Name UserPrincipalName1 -Value $_.TargetUserPrincipalName -Force
  
  $MappingFile += $thisMappingObject

  #Collection Object
  $thisCollectionObject = [PSCustomObject]@{
    UserPrincipalName = $_.SourceUserPrincipalName
    Collection        = $CollectionName
  }

  $CollectionFile += $thisCollectionObject

}

$exportableNames = "$($TODAY)_$($CollectionName)"

if ($GenerateFromIAMFile -eq $true) {
  $folderPath = $IAMRepositoryPath
  Write-Host "IAMREPO: $($folderPath)" -ForegroundColor Yellow
}
elseif ($GenerateFromIAMFile -eq $false) {
  $folderPath = "$($PATH_FILES)\$($exportableNames)"
  
  if (( Test-Path -Path $folderPath) -eq $false) {
    Write-Host "Creating folder"
    New-Item -ItemType Directory -Path $folderPath
  }
}



Write-Host "Exporting Files"
$DiscoverFile | Export-Csv -Path "$($folderPath)\$($exportableNames)_Discovery.csv" -NoTypeInformation
$MappingFile | Export-Csv -Path "$($folderPath)\$($exportableNames)_Mapping.csv" -NoTypeInformation
(Get-Content -Path "$($folderPath)\$($exportableNames)_Mapping.csv" -Raw) -replace 'UserPrincipalName1', 'UserPrincipalName' | Set-Content -Path "$($folderPath)\$($exportableNames)_Mapping.csv"
$CollectionFile | Export-Csv -Path "$($folderPath)\$($exportableNames)_Collection.csv" -NoTypeInformation

