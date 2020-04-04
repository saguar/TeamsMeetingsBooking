[CmdletBinding()]
param (
	[Parameter()]
	[string]
	$webSiteName,
	[Parameter()]
	[string]
	$objectID
)
Connect-AzureAD
$webSiteServicePrincipal = Get-AzureADServicePrincipal -Filter "displayName eq '$webSiteName'" | Where-Object ObjectId -eq "$objectID"
$graphServicePrincipal = Get-AzureADServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
$appRole = $graphServicePrincipal.AppRoles | Where-Object {$_.Value -eq 'Calendars.ReadWrite' -and $_.AllowedMemberTypes -contains 'Application'}    

New-AzureAdServiceAppRoleAssignment `
	-ObjectId $webSiteServicePrincipal.ObjectId `
	-PrincipalId $webSiteServicePrincipal.ObjectId `
	-ResourceId $graphServicePrincipal.ObjectId `
	-Id $appRole.Id

Get-AzureADServiceAppRoleAssignment `
-ObjectId $graphServicePrincipal.ObjectId `
| Where-Object { $_.Id -eq $appRole.Id -and $_.PrincipalId -eq
	$webSiteServicePrincipal.ObjectId }

$appRole = $graphServicePrincipal.AppRoles | Where-Object {$_.Value -eq 'OnlineMeetings.ReadWrite.All' -and $_.AllowedMemberTypes -contains 'Application'}    

New-AzureAdServiceAppRoleAssignment `
	-ObjectId $webSiteServicePrincipal.ObjectId `
	-PrincipalId $webSiteServicePrincipal.ObjectId `
	-ResourceId $graphServicePrincipal.ObjectId `
	-Id $appRole.Id

	Get-AzureADServiceAppRoleAssignment `
	-ObjectId $graphServicePrincipal.ObjectId `
	| Where-Object { $_.Id -eq $appRole.Id -and $_.PrincipalId -eq
		$webSiteServicePrincipal.ObjectId }
	
$appRole = $graphServicePrincipal.AppRoles | Where-Object {$_.Value -eq 'User.Read.All' -and $_.AllowedMemberTypes -contains 'Application'}    

New-AzureAdServiceAppRoleAssignment `
	-ObjectId $webSiteServicePrincipal.ObjectId `
	-PrincipalId $webSiteServicePrincipal.ObjectId `
	-ResourceId $graphServicePrincipal.ObjectId `
	-Id $appRole.Id

Get-AzureADServiceAppRoleAssignment `
-ObjectId $graphServicePrincipal.ObjectId `
| Where-Object { $_.Id -eq $appRole.Id -and $_.PrincipalId -eq
	$webSiteServicePrincipal.ObjectId }