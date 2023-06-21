<#

.SYNOPSIS
PowerShell script to Connect to Microsoft Teams
The advantage of this script over the standard 'Connect-MicrosoftTeams' command is that you get feedback about the tenant you are connected to. This is useful for companies manging multiple tenants.
By Ruud van Strijp - Axians
ruud.vanstrijp@axians.com

.NOTES
Microsoft Teams module 4.0.0 or higher is highly recommended, things might break if you run an older version
Method 1 (Recommended)
Update-Module MicrosoftTeams

Method 2 (Alternative)
Uninstall-Module -Name MicrosoftTeams -AllVersions
Install-Module -Name MicrosoftTeams -Force -Scope AllUsers
#>

Import-module MicrosoftTeams

$teamsModuleVersion = (Get-InstalledModule -Name MicrosoftTeams).Version
if($teamsModuleVersion -lt 4.0.0){
    Write-Host "  WARNING: Module Version older than 4.0.0 will be deprecated soon. This script might not run well" -ForegroundColor red
}

Connect-MicrosoftTeams

$tenant = Get-CsTenant

Write-Host "  Connected to tenant: " -ForegroundColor White -NoNewLine
Write-Host $tenant.DisplayName -ForegroundColor Green
Write-Host "  Tenant city: " -ForegroundColor White -NoNewLine
Write-Host $tenant.City -ForegroundColor Green
Write-Host "  Service instance: " -ForegroundColor White -NoNewLine
$instance = $tenant.ServiceInstance -replace ".*/"
Write-Host $instance -ForegroundColor Green
write-host ""
<#
$tenantIdentity = (Get-CsTenant).Identity -replace ".*lync","" -replace "001.*",""
$adminURL = "https://admin$($tenantIdentity).online.lync.com/HostedMigration/hostedmigrationService.svc"

Write-Host "  Admin URL: " -ForegroundColor White -NoNewLine
Write-Host $adminURL -ForegroundColor Green
#>

Write-Host "  SIP domains: " -ForegroundColor White
$sipDomains = $tenant.SipDomain
foreach ($sipDomain in $sipDomains) {
   Write-Host " "$sipDomain -ForegroundColor Yellow
}
write-host ""