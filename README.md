# GraphToolkit
Microsoft Graph Toolkit

Use Get-AutRefreshToken to request a token from Microsoft Graph API.
Needed to perform this task are the following 3 values:
AppId 
AppSecret
TenantId

These values can be found on your application in Azure.

$Token = Get-AutRefreshToken -AppId $AppId -appSecret $AppSecret -tenantId $TenantId

Example on how to query Graph Api for all Users.
Connect-AutGraph -Endpoint "Users" -Access_token $Token -All Yes

