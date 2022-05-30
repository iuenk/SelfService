<#PSScriptInfo
.VERSION 1.0
.AUTHOR Ivo Uenk
.RELEASENOTES

#>
<#
.SYNOPSIS
  Remove user via PowerAutomate
.DESCRIPTION
  Remove user via PowerAutomate
.NOTES
  Version:        1.0
  Author:         Ivo Uenk
  Creation Date:  2022-03-30
  Purpose/Change: Remove user via PowerAutomate

  Install the following modules in Automation Accounts:
  AzureADPreview
  ExchangeOnlineManagement

#>

# Parameters from Power Automate
Param (
  [string] $UserPrincipalName,
  [array] $Delegates
)

# Import the necessary modules
Import-Module 'AzureADPreview'
Import-Module 'ExchangeOnlineManagement'
Import-Module 'Microsoft.Online.SharePoint.PowerShell'

# Global variables
$LicenseName = "LIC-CL-"

# Get the credential from Automation  
$credential = Get-AutomationPSCredential -Name 'AutomationCreds'  
$userName = $credential.UserName  
$securePassword = $credential.Password
$psCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userName, $securePassword

# Connect to Microosft services
Connect-AzureAD -Credential $psCredential
Connect-ExchangeOnline -Credential $psCredential

$FirstName = ((Get-AzureADUser -ObjectId $UserPrincipalName).GivenName)
$LastName = ((Get-AzureADUser -ObjectId $UserPrincipalName).Surname)
$Date = Get-Date -Format "dddd, dd/MM/yyyy"

# Rename and disable user account
$ExitDate = (Get-Date -Format "ddMMyyyy").Insert(0,'Exit')
$DisplayName = $ExitDate + " " + $FirstName + " " + $LastName

Set-AzureADUser -ObjectId $UserPrincipalName -DisplayName $DisplayName -AccountEnabled $false
Write-Output "Disable user $UserPrincipalName on $Date"

#Mailbox actions
# Check if user mailbox exist
$Mailbox = ((Get-AzureADUser -ObjectId $UserPrincipalName).mail)

If (!(Get-Mailbox $Mailbox))
{
  Write-Output "No user mailbox for $UserPrincipalName found"

} else {
# Check if mailbox is larger than 50GB and convert it to shared mailbox
$Stats = (Get-MailboxStatistics -Identity $Mailbox | Select-Object DisplayName, @{Name="TotalItemSizeMB"; Expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),0)}})
$Size = $Stats.TotalItemSizeMB
Write-Output "Mailbox size for $Mailbox is $Size MB"

    if ($Size -ge 50000)
    {
        Write-Output "Result:Failed (Mailbox size equals or greather than 50GB)"
        Write-Output "Description: Mailbox for user $Mailbox is to large to convert to shared mailbox"
        $_.Exception.Message
        $_.InvocationInfo.ScriptLineNumber
        exit

    } else {
    
    # Remove all calendar meetings for the mailbox
    Remove-CalendarEvents -Identity $Mailbox -CancelOrganizedMeetings -Confirm:$False -QueryWindowInDays 1825
    Write-Output "Remove calendar events for $Mailbox"
    
    # Convert mailbox to Shared mailbox
    Set-Mailbox -Identity $Mailbox -Type Shared
    Write-Output "Convert mailbox $Mailbox to shared mailbox"
    
    # Estimated time to complete converting the mailbox is 5 minutes
    Start-Sleep -Seconds 300

    # Give delegate(s) permissions on mailbox and set forwarder
    foreach($Delegate in $Delegates)
    {
    $DelegateUPN = $Delegate.UserPrincipalName  
    # Set forwarder on mailbox
    Set-Mailbox -Identity $Mailbox -DeliverToMailboxAndForward $false -ForwardingSMTPAddress $DelegateUPN
    Write-Output "Set forwarder $DelegateUPN for shared mailbox $Mailbox"

    # Set permissions on mailbox
    Add-MailboxPermission -Identity $Mailbox -User $DelegateUPN -AccessRights FullAccess -InheritanceType All -AutoMapping $false
    Write-Output "Give $DelegateUPN full access on shared mailbox $Mailbox"
    }
  }
} 

# Remove user from license groups, mailbox and onedrive will be removed!
$Licenses = Get-AzureADUserLicenseDetail -ObjectId $UserPrincipalName
$AssignedUser = Get-AzureADUser | Where-Object { $_.UserPrincipalName -eq "$UserPrincipalName"}

if ($Licenses)
{
  $LicenseGroups = Get-AzureADGroup -SearchString $LicenseName

  Foreach ($LicenseGroup in $LicenseGroups)
  {
    Try 
    {
        # Remove all license groups from user
        remove-AzureADGroupMember -ObjectId $LicenseGroup.ObjectId -MemberId $AssignedUser.ObjectId
        Write-Output "Remove license group $($LicenseGroup.DisplayName) for $UserPrincipalName"

    } Catch {
        Write-Output "$UserPrincipalName already removed from $($LicenseGroup.DisplayName)"
    }
  }
}