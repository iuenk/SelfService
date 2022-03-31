Install-Module -Name Microsoft.PowerApps.Administration.PowerShell
Install-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber

# Here is how you can pass in credentials (avoiding opening a prompt)
$pass = ConvertTo-SecureString "<Password>" -AsPlainText -Force
Add-PowerAppsAccount -Username <automation account> -Password $pass

# Get environment name
Get-AdminPowerAppEnvironment | select EnvironmentName

# App ID <App ID> (Self-Service Ucorp)
Set-AdminPowerAppApisToBypassConsent -AppName <App ID> -EnvironmentName Default-<EnvironmentName>