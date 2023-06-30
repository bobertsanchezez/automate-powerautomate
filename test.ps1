Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser
Install-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber -Scope CurrentUser
# This call opens prompt to collect credentials (Azure Active Directory account and password) used by the commands 
Add-PowerAppsAccount -Endpoint "usgov"

# Get flows shared with you in a specific environment
Get-AdminFlow -Verbose
# Display information about the shared flows
# $flows | Format-Table Name, Owner, Enabled
