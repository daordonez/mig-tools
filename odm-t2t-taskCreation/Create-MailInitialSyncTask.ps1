param(
  [Parameter(Mandatory = $true)]$CollectionName
)

#Const
$__CONF = (Get-Content -Path '.\project_config.json') | ConvertFrom-Json
$TASK_CONF = $__CONF.params.projectConf.tasks.mailTasks.contentMigration

#Task params definition
$TASK_PREFIX = $TASK_CONF.preMigration.prefixName
$TASK_SUFIX = $TASK_CONF.commonNameSufix
$TASK_MAIL_FWD = $TASK_CONF.forwardingConfig
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
  $thisMailMigrationTask = New-OdmMailMigrationTask -Name $TaskName -MigrateMail -MigrateRecoverableItems -MailForwarding $TASK_MAIL_FWD  `
    -O365LicenseAssignmentType $TASK_LIC_ASSIGNMENT_TYPE -O365LicenseToAssign $TASK_LIC_ASSIGNMENT

  #Bypass license bug assingment ODM
  Set-OdmMailMigrationTask -Task $thisMailMigrationTask -O365LicenseToAssign ""

  #Check creation status
  if ($null -ne $thisMailMigrationTask) {
    Write-Host "Task: $($TaskName) succesfully created" -ForegroundColor Green

    #Add objects to task
    Write-Host "Adding Objects to task"

    Add-OdmObject -To $thisMailMigrationTask -Objects $allUsers

    #Output Task result
    Get-OdmTask -Task $thisMailMigrationTask

    Write-Host "Script end"

  }
  else {
    Write-Host "Task: $($TaskName) creation failed" -ForegroundColor Red
  }


}
else {
  Write-Host "No collection found with name:$($CollectionName)" -ForegroundColor Yellow
}

