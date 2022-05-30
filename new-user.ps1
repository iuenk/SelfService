<#PSScriptInfo
.VERSION 1.0
.AUTHOR Ivo Uenk
.RELEASENOTES

#>
<#
.SYNOPSIS
  New user via PowerAutomate
.DESCRIPTION
  New user via PowerAutomate
.NOTES
  Version:        1.0
  Author:         Ivo Uenk
  Creation Date:  2022-03-30
  Purpose/Change: New user via PowerAutomate

  Install the following modules in Automation Accounts:
  AzureADPreview

#>

# Parameters from Power Automate
Param (
  [string] $FirstName,
  [string] $LastName,
  [string] $Department,
  [string] $Other,
  [string] $MobilePhone,
  [string] $Password,
  [string] $UserPrincipalName,
  [array] $Licenses,
  [boolean] $AccountStatus
)

# Import the necessary modules
Import-Module 'AzureADPreview'

# Global variables
$Date = Get-Date -Format "dddd, dd/MM/yyyy"
$DisplayName = $FirstName + " " + $LastName
$MailNickName = $FirstName + $LastName
$UsageLocation = "NL"
$LicenseName = "LIC-CL-"

# Get the credential from Automation  
$credential = Get-AutomationPSCredential -Name 'AutomationCreds'  
$userName = $credential.UserName  
$securePassword = $credential.Password
$psCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userName, $securePassword

# Generate password profile
$PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
$PasswordProfile.Password = $Password
$PasswordProfile.ForceChangePasswordNextLogin = $true

# Connect to Microsoft services
Connect-AzureAD -Credential $psCredential

# Check available licenses
$NeededLicenses = @()

# Loop trough all license groups and format those to SkuPartNumber
# The Azure AD group LIC-CL-NOLICENSE is used as dummy group
foreach ($License in $Licenses)
{
	$neededlicense = $License.displayName
	$neededlicense = $neededlicense.Replace($LicenseName, "")
	$neededlicense = $neededlicense.Split("-")
	$NeededLicenses += $neededlicense
}

$NeededLicenses = $NeededLicenses | Select-Object -Unique

if ($NeededLicenses -notcontains "NOLICENSE")
{
    # Check if the licenses needed are available, if not, stop the script otherwise create user.
    foreach ($Neededlicense in $NeededLicenses)
    {
	    Write-Output "Check for license $Neededlicense"
	    $availablelicenses = Get-AzureADSubscribedSku | Select-Object SkuPartNumber,ConsumedUnits -ExpandProperty Prepaidunits| Where-Object {$_.SkuPartNumber -eq $Neededlicense}
	    $countavailablelicenses = $availablelicenses.Enabled - $availablelicenses.ConsumedUnits
	    Write-Output "${countavailablelicenses}: $Neededlicense"
	    if ($countavailablelicenses -lt 1)
	    {
		    Write-Output "Result:Failed (not enough licenses)"
		    Write-Output "Description: LicensesNeeded for $Neededlicense"
		    $_.Exception.Message
		    $_.InvocationInfo.ScriptLineNumber
		    exit
	    }
    }

} else {
    Write-Output "No license is needed for $UserPrincipalName stop counting available licenses"
} 

# Create new user
Write-Output "Creating user $UserPrincipalName on $Date"

New-AzureADUser -DisplayName $DisplayName `
				-PasswordProfile $PasswordProfile `
				-GivenName $FirstName -Surname $LastName `
				-UserPrincipalName $UserPrincipalName `
				-Department $Department `
				-Mobile $MobilePhone `
				-MailNickName $MailNickName `
				-UsageLocation $UsageLocation `
				-AccountEnabled $AccountStatus

# Set extra department attribute
if ($Other)
{
	# Retrieve extension property
	$MyApp = (Get-AzureADApplication -SearchString "Custom Extensions").objectId
	$E = Get-AzureADApplicationExtensionProperty -ObjectId $MyApp
	$Extension = $E.Name
	
	Set-AzureADUserExtension -ObjectId $UserPrincipalName -ExtensionName $Extension -ExtensionValue $Other
}

# Assign license to user via license groups
if ($NeededLicenses -notcontains "NOLICENSE")
{
	foreach ($License in $Licenses)
	{
		$LicenseGroup = $License.displayName
		$AssignedGroup = Get-AzureADGroup -Filter "DisplayName eq '$LicenseGroup'"
		$AssignedUser = Get-AzureADUser | Where-Object { $_.UserPrincipalName -eq "$UserPrincipalName"}
		Add-AzureADGroupMember -ObjectId $AssignedGroup.ObjectId -RefObjectId $AssignedUser.ObjectId
		Write-Output "$UserPrincipalName added to group $LicenseGroup"
	}

} Else {
	Write-Output "Do not assign license(s) to $UserPrincipalName"
}