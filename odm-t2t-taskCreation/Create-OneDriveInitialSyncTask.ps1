param(
  [Parameter(Mandatory = $true)]$CollectionName
)

#Const
$__CONF = (Get-Content -Path '.\project_config.json') | ConvertFrom-Json
$TASK_CONF = $__CONF.params.projectConf.tasks.oneDriveTasks.contentMigration

#Task params definition
$TASK_PREFIX = $TASK_CONF.preMigration.prefixName
$TASK_SUFIX = $TASK_CONF.commonNameSufix
$TASK_LIC_ASSIGNMENT_TYPE = $TASK_CONF.licenses.O365LicenseAssignmentType
$TASK_LIC_ASSIGNMENT = $TASK_CONF.licenses.O365LicenseToAssign
$TODAY = Get-Date

#Get collection
Write-Host "Getting collection:$($CollectionName)"
$collectionName = @{'name' = $CollectionName }
$thisCollection = Get-OdmCollection -WildcardFilter $collectionName

#Check if collections exists
if ($null -ne $thisCollection) {
  #Get all collection members
  Write-Host "Getting collection members"
  $allUsers = @()
  $thisUsersCol = $thisCollection | Get-OdmObject 
  $thisUsersCol | ForEach-Object {
    $filter = @{'sourceEmail' = $_.SourceEmail }
    $thisUserToTask = Get-OdmObject -WildcardFilter $filter
    $allUsers += $thisUserToTask
  }

  #Output current collection members
  $allUsers | Format-Table sourceEmail, targetEmail, mailMigrationStatus -AutoSize

  #Create new mail migration task
  
  $toDayToString = $TODAY | Get-Date -Format "yyyyMMdd"
  $TaskName = "$($toDayToString)_$($TASK_PREFIX)_$($TASK_SUFIX)"

  Write-Host "Creating task: $($TaskName) "
  $thisODMigrationTask = New-OdmOneDriveMigrationTask -Name $TaskName -MigrationAction Skip -FileVersions LatestAndPrevious `
  -PermissionBehaviour MigratedContentOnly -Author TargetAccount -FileVersionMaxSize 32 -O365LicenseAssignmentType $TASK_LIC_ASSIGNMENT_TYPE -O365LicenseToAssign $TASK_LIC_ASSIGNMENT

  #Bypass license bug assingment ODM
  Set-OdmOneDriveMigrationTask -Task $thisODMigrationTask -O365LicenseToAssign ""

  #Check creation status
  if ($null -ne $thisODMigrationTask) {
    Write-Host "Task: $($TaskName) succesfully created" -ForegroundColor Green

    #Add objects to task
    Write-Host "Adding Objects to task"

    Add-OdmObject -To $thisODMigrationTask -Objects $allUsers

    #Output Task result
    Get-OdmTask -Task $thisODMigrationTask 

    Write-Host "Script end"

  }
  else {
    Write-Host "Task: $($TaskName) creation failed" -ForegroundColor Red
  }


}
else {
  Write-Host "No collection found with name:$($CollectionName)" -ForegroundColor Yellow
}

