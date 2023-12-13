
<#
  Date: 13/12/2022
  Author: Diego Ordonez
  Sinopsis: Quest On Demand migration API connection. 
#>


#CONST
$__CONF = (Get-Content -Path '.\project_config.json') | ConvertFrom-Json
$ODM_REGION = $__CONF.params.tenantConfiguration.region
$ODM_ORGID = $__CONF.params.tenantConfiguration.organizationId
$ODM_PROJECTID = $__CONF.params.tenantConfiguration.projectId

#load ODM API
Import-Module $__CONF.apiPath -Global

. "C:\Program Files\WindowsPowerShell\Modules\ODMApi\OdmAPI.Types.ps1"

#Service connection
Connect-OdmService -Region $ODM_REGION

#Organization
Select-OdmOrganization -OrganizationId $ODM_ORGID


#Project selection
Select-OdmProject $ODM_PROJECTID

$projectName = (Get-OdmProject -Project $ODM_PROJECTID).Name


if( -not ( $null -eq $projectName )){
  Write-Host "ODM Conection successfully. Project Name: $($projectName)" -ForegroundColor Green
}else {
  Write-Host "Error. Connection could not be stablished" -ForegroundColor Red
}
