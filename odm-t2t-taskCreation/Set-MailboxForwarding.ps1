param(
  [Parameter(Mandatory = $true)]$CollectionName,
  [Parameter(Mandatory = $true)][ValidateSet("Initial", "Final")]$Phase
)

#Const
$__CONF = (Get-Content -Path '.\project_config.json') | ConvertFrom-Json
$TASK_CONF = $__CONF.params.projectConf.tasks.mailTasks.mailForwarding

#Task params definition
$TASK_SUFIX = $TASK_CONF.commonNameSufix
$HAS_INITIAL_FWD = $TASK_CONF.mailForwarding.forwardingTask.initialSwitchDirectionEnable
$ALLOW_EXEC = $false
$TODAY = Get-Date
$TASK_FWD_DIRECTION = ""

switch ($Phase) {
  Initial {
    $TASK_PREFIX = $TASK_CONF.preMigration.prefixName
    $TASK_FWD_DIRECTION = $TASK_CONF.preMigration.switchDirection
  }
  Final {
    $TASK_PREFIX = $TASK_CONF.migration.prefixName
    $TASK_FWD_DIRECTION = $TASK_CONF.migration.switchDirection
  }
}

#Create new mail forward task
function New-ForwardTask {

  $toDayToString = $TODAY | Get-Date -Format "yyyyMMdd"
  $TaskName = "$($toDayToString)_$($TASK_PREFIX)_$($TASK_SUFIX)"
  
  Write-Host "Creating task: $($TaskName) "
  $thisMailForwardTask = New-OdmMailSwitchTask -Name $TaskName -SwitchDirection $TASK_FWD_DIRECTION

  return $thisMailForwardTask
}


if ($HAS_INITIAL_FWD -like "True") {
  $ALLOW_EXEC = $true
}

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

  if (($Phase -eq "Initial") -and ($ALLOW_EXEC -eq $true)) {
  }
  else {
  }
  switch ($Phase) {
    Initial {
      if($ALLOW_EXEC -eq $true){
        $thisMailForwardTask = New-ForwardTask
      }else{
        Write-Host "Script end. 'InitialSwitchDirectionEnable' is disabled for this project. Please review 'project_config.json' for allow execution" -ForegroundColor Yellow 
      }
    }
    Final {
      $thisMailForwardTask = New-ForwardTask
    }
  }
  
  


  #Check creation status
  if ($null -ne $thisMailForwardTask) {
    Write-Host "Task: $($TaskName) succesfully created" -ForegroundColor Green

    #Add objects to task
    Write-Host "Adding Objects to task"

    Add-OdmObject -To $thisMailForwardTask -Objects $allUsers

    #Output Task result
    Get-OdmTask -Task $thisMailForwardTask

    Write-Host "Script end"

  }
  else {
    Write-Host "Task: $($TaskName) creation failed" -ForegroundColor Red
  }


}
else {
  Write-Host "No collection found with name:$($CollectionName)" -ForegroundColor Yellow
}





