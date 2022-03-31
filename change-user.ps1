<#PSScriptInfo
.VERSION 1.0
.AUTHOR Ivo Uenk
.RELEASENOTES

#>
<#
.SYNOPSIS
  Change user via PowerAutomate
.DESCRIPTION
  Change user via PowerAutomate
.NOTES
  Version:        1.0
  Author:         Ivo Uenk
  Creation Date:  2022-03-30
  Purpose/Change: Change user via PowerAutomate

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
  [string] $UserPrincipalName,
  [array] $Licenses,
  [boolean] $AccountStatus
)

# Import the necessary modules
Import-Module 'AzureADPreview'

# Variables
# Get the credential from Automation  
$credential = Get-AutomationPSCredential -Name 'AutomationCreds'  
$userName = $credential.UserName  
$securePassword = $credential.Password

$PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
$PasswordProfile.Password = $Password
$PasswordProfile.ForceChangePasswordNextLogin = $true

# Get Office 365 credential from Azure Automation and connect to Azure AD
$psCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userName, $securePassword
Connect-AzureAD -Credential $psCredential

$DisplayName = $FirstName + " " + $LastName
$LicenseName = "LIC-CL-"

$Date = Get-Date -Format "dddd, dd/MM/yyyy"

# Change user data
Write-Output "Changing user $UserPrincipalName on $Date"

Set-AzureADUser -ObjectId $UserPrincipalName `
                -DisplayName $DisplayName `
                -GivenName $FirstName -Surname $LastName `
                -Department $Department `
                -Mobile $MobilePhone `
                -AccountEnabled $AccountStatus

# Retrieve extension property
$MyApp = (Get-AzureADApplication -SearchString "Custom Extensions").objectId
$E = Get-AzureADApplicationExtensionProperty -ObjectId $MyApp
$Extension = $E.Name

# Set or remove extra department attribute
if ($Other)
{
	Set-AzureADUserExtension -ObjectId $UserPrincipalName -ExtensionName $Extension -ExtensionValue $Other
  Write-Output "Changing user $UserPrincipalName other department to $Other on $Date"

} else {
    Remove-AzureADUserExtension -ObjectId $UserPrincipalName -ExtensionName $Extension
    Write-Output "Remove user $UserPrincipalName other department on $Date"
}

# Check if the licenses need to be added or not
$AssignedUser = Get-AzureADUser | Where-Object { $_.UserPrincipalName -eq "$UserPrincipalName"}

# Check available licenses
$NeededLicenses = @()

# Loop trough all licenses
foreach ($License in $Licenses)
{
  $neededlicense = $License.displayName
  $neededlicense = $neededlicense.Replace($LicenseName, "")
  $neededlicense = $neededlicense.Split("-")
  $NeededLicenses += $neededlicense
}

$NeededLicenses = $NeededLicenses | Select-Object -Unique
$AssignedLicenses = Get-AzureADUserLicenseDetail -ObjectId $AssignedUser.objectid | Select-Object -ExpandProperty SkuPartNumber

if ($NeededLicenses -notcontains "NOLICENSE")
{
	foreach ($Neededlicense in $Neededlicenses)
  {
    if ($Neededlicense -inotin $AssignedLicenses)
    {
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
        
      } else {
          $LicenseGroup = $License.displayName
          $AssignedGroup = Get-AzureADGroup -Filter "DisplayName eq '$LicenseGroup'"
          Add-AzureADGroupMember -ObjectId $AssignedGroup.ObjectId -RefObjectId $AssignedUser.ObjectId
          Write-Output "$UserPrincipalName added to group $LicenseGroup"       
      }
    }
  }

} else {
  Write-Output "No license is needed for $UserPrincipalName stop counting available licenses"
}

# Assign license to user via license groups
if ($NeededLicenses -notcontains "NOLICENSE")
{
	foreach ($License in $Licenses)
	{
		$LicenseGroup = $License.displayName
		$AssignedGroup = Get-AzureADGroup -Filter "DisplayName eq '$LicenseGroup'"
		$AssignedUser = Get-AzureADUser | Where-Object { $_.UserPrincipalName -eq "$UserPrincipalName"}

        Try 
        {
		    Add-AzureADGroupMember -ObjectId $AssignedGroup.ObjectId -RefObjectId $AssignedUser.ObjectId
		    Write-Output "$UserPrincipalName added to group $LicenseGroup"

        } Catch {
            Write-Output "$UserPrincipalName already present in $($AssignedGroup.DisplayName)"
	    }
  }

} Else {
# Remove user from license groups, mailbox and onedrive will be removed!
  Foreach ($License in $Licenses)
  {
    $LicenseGroup = $License.displayName
    $AssignedGroup = Get-AzureADGroup -Filter "DisplayName eq '$LicenseGroup'"
    $AssignedUser = Get-AzureADUser | Where-Object { $_.UserPrincipalName -eq "$UserPrincipalName"}

      Try 
      {
          remove-AzureADGroupMember -ObjectId $LicenseGroup.ObjectId -MemberId $AssignedUser.ObjectId
          Write-Output "Remove license group $($LicenseGroup.DisplayName) for $UserPrincipalName"
          
      } Catch {
          Write-Output "$UserPrincipalName already removed from $($LicenseGroup.DisplayName)"
    }
  }
}