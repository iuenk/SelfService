  <#PSScriptInfo
.VERSION 1.0
.AUTHOR Ivo Uenk
.RELEASENOTES

#>
<#
.SYNOPSIS
  Change group via PowerAutomate
.DESCRIPTION
  Change group via PowerAutomate
.NOTES
  Version:        1.0
  Author:         Ivo Uenk
  Creation Date:  2022-03-30
  Purpose/Change: Change group via PowerAutomate

  Install the following modules in Automation Accounts:
  AzureADPreview
  Microsoft.Online.SharePoint.PowerShell
  MicrosoftTeams
  ExchangeOnlineManagement

#>

# Parameters from Power Automate
Param (
  [string] $GroupName,
  [string] $Department,
  [string] $AccessType,
  [string] $ExternalSharing,
  [boolean] $GuestAccess,
  [boolean] $Teams
)

# Import the necessary modules
Import-Module 'AzureADPreview'
Import-Module 'Microsoft.Online.SharePoint.PowerShell'
Import-Module 'MicrosoftTeams'
Import-Module 'ExchangeOnlineManagement'

# Global variables
$Tenant = Get-AutomationVariable -Name 'Tenant'

# Get the credential from Automation  
$credential = Get-AutomationPSCredential -Name 'AutomationCreds'  
$userName = $credential.UserName  
$securePassword = $credential.Password

$PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
$PasswordProfile.Password = $Password
$PasswordProfile.ForceChangePasswordNextLogin = $true
$psCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userName, $securePassword

# Connect to Microsoft 365 services
Connect-AzureAD -Credential $psCredential
Connect-ExchangeOnline -Credential $psCredential
Connect-SPOService -Url "https://$($Tenant)-admin.sharepoint.com" -Credential $psCredential
Connect-MicrosoftTeams -Credential $psCredential -ErrorAction Stop  

# Microsoft 365 variables
$Alias = $GroupName.Replace(' ','').Replace('&','').ToLower()

# Teams variables
$AllowCreateUpdateRemoveTabs = $false
$AllowAddRemoveApps =$false
$AllowDeleteChannels =$false
$AllowCreateUpdateChannels =$false
$AllowCreateUpdateRemoveConnectors =$false
$SiteURL = "https://$($Tenant).sharepoint.com/sites/$Alias"
$Site = Get-SPOSite -Identity $SiteURL
$TeamsExists = Get-Team -DisplayName $GroupName -ErrorAction Stop

# Will be used for dynamic groups as used in the GroupTypes property of a group object
$dynamicGroupTypeString = "DynamicMembership"

function ConvertStaticGroupToDynamic
{
    Param([string]$groupId, [string]$dynamicMembershipRule)
    $dynamicGroupTypeString = "DynamicMembership"
    #existing group types
    [System.Collections.ArrayList]$groupTypes = (Get-AzureAdMsGroup -Id $groupId).GroupTypes

    #add the type for dynamic groups
    $groupTypes.Add($dynamicGroupTypeString)
    $groupTypes = $groupTypes | Select-Object -Unique

    # Convert static group to dynamic group
    Set-AzureAdMsGroup -Id $groupId -GroupTypes $groupTypes.ToArray() -MembershipRuleProcessingState "On" -MembershipRule $dynamicMembershipRule
}

function ConvertDynamicGroupToStatic
{
    Param([string]$groupId)

    #existing group types
    [System.Collections.ArrayList]$groupTypes = (Get-AzureAdMsGroup -Id $groupId).GroupTypes

    #remove the type for dynamic groups, but keep the other type values
    $groupTypes.Remove($dynamicGroupTypeString)

    #modify the group properties to make it a static group: i) change GroupTypes to remove the dynamic type, ii) pause execution of the current rule
    Set-AzureAdMsGroup -Id $groupId -GroupTypes $groupTypes.ToArray()
}

# Create Microsoft 365 group
if ($Department)
{
    # Retrieve Azure AD extension property
    $Identity = (Get-UnifiedGroup -Identity $GroupName)
    $MyApp = (Get-AzureADApplication -SearchString "Custom Extensions").objectId
    $E = Get-AzureADApplicationExtensionProperty -ObjectId $MyApp
    $Extension = $E.Name.Insert(0,'user.')

    $ID = (Get-UnifiedGroup $Identity.Name).ExternalDirectoryObjectId
    ConvertStaticGroupToDynamic -groupId $ID -dynamicMembershipRule "(user.department -eq ""$Department"") or ($Extension -eq ""$Department"") or (user.department -eq ""Board"")"
    Write-Output "Assigned members that do not have department $Department will be removed from $GroupName"
    Write-Output "Convert static group $GroupName to dynamic"

} Else {
    $Identity = (Get-UnifiedGroup -Identity $GroupName)
    $ID = (Get-UnifiedGroup $Identity.Name).ExternalDirectoryObjectId
    ConvertDynamicGroupToStatic -groupId $ID
    Write-Output "All members of dynamic group $GroupName will now be assigned member"
    Write-Output "Convert dynamic group $GroupName to static"
}

# Set Access type public or private
Set-UnifiedGroup -Identity $GroupName -AccessType $AccessType
Write-Output "Changed UnifiedGroup '$GroupName'"

# Set External sharing Disabled, ExistingExternaluserSharingOnly, ExternaluserSharingOnly
if ($ExternalSharing)
{
    Write-Output "Set external sharing $ExternalSharing for $GroupName"
    Set-SPOSite -Identity "https://$tenant.sharepoint.com/sites/$Alias" -SharingCapability $ExternalSharing -ErrorAction Stop | Out-Null  
}

# Create visitors group  
if($GuestAccess -eq $true)
{
    $GuestGroup = Get-AzureADGroup -SearchString "sp-visitors-$Alias" -ErrorAction stop
    if($Null -eq $GuestGroup)
    {
        New-AzureADGroup -DisplayName "sp-visitors-$Alias" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet" -Description "Read access to SharePoint site `"$GroupName`""
        Write-Output "SharePoint read only group 'sp-visitors-$Alias' is created" 

    } Else {
    Write-Output "SharePoint read only group 'sp-visitors-$Alias' already exists"
    }
}

If($GuestAccess -eq $true)
{
    $SPOGroup = "$GroupName Visitors"
    $LoginName = "sp-visitors-$Alias"
    [bool]$Success = $false
    Do
    {
        try
        {
            Add-SPOUser -Group $SPOGroup -LoginName $LoginName -Site "https://$($Tenant).sharepoint.com/sites/$Alias" -ErrorAction Stop
            Write-Output "Group '$LoginName' is added to site $Alias" 
            [bool]$Success = $true

        } catch {
            [bool]$Success=$false
            Write-Output "Wait 60 seconds and try again to add $LoginName"
            Start-Sleep 60
        }
    } until($Success)
}

# Remove visitors group
If($GuestAccess -eq $false)
{
    $SPOGroup = "$GroupName Visitors"
    $Alias = $GroupName.Replace(' ','').Replace('&','').ToLower()
    $LoginName = "sp-visitors-$Alias"

    Remove-SPOUser -Group $SPOGroup -LoginName $LoginName -Site "https://$($Tenant).sharepoint.com/sites/$Alias" -ErrorAction Stop
    Write-Output "Group '$LoginName' is removed from site $Alias'" 
}

if($GuestAccess -eq $false)
{
    $GuestGroups = Get-AzureADGroup -SearchString "sp-visitors-$Alias" -ErrorAction stop
    Foreach ($GuestGroup in $GuestGroups)
    {
        Remove-AzureADGroup -ObjectId $GuestGroup.ObjectId
        Write-Output "SharePoint read only group sp-visitors-$Alias is removed"
    }
}

# Enable Teams functionality for SPO Site
If ($Teams -eq $true)
{
    if ($Null -eq $TeamsExists) 
    {
        $NewTeams = New-Team -GroupId $Site.GroupId.Guid -ErrorAction Stop
        Set-Team -GroupId $NewTeams.GroupId `
            -AllowCreateUpdateRemoveConnectors $AllowCreateUpdateRemoveConnectors `
            -AllowCreateUpdateChannels $AllowCreateUpdateChannels `
            -AllowDeleteChannels $AllowDeleteChannels `
            -AllowAddRemoveApps $AllowAddRemoveApps `
            -AllowCreateUpdateRemoveTabs $AllowCreateUpdateRemoveTabs `
            -ErrorAction Stop | Out-Null

        Write-Output "Teams enabled for site $GroupName"

    } Else {
        $Archived = Get-Team -GroupId $Site.GroupId.Guid | Select-Object archived
        
        if ($Archived.archived -eq $true)
        {
            Set-TeamArchivedState -GroupId $Site.GroupId.Guid -Archived:$False -ErrorAction SilentlyContinue | Out-Null
            Write-Output "Teams already exist for site $GroupName and archiving is turned off"

        } Else {
            Write-Output "Teams already exist for site $GroupName and is not archived"
        }
    }
}

# Disable Teams functionality for SPO Site
If ($Teams -eq $false)
{
    If ($TeamsExists)
    {
        $Archived = Get-Team -GroupId $Site.GroupId.Guid | Select-Object archived
        if ($Archived.archived -eq $false)
        {
            Set-TeamArchivedState -GroupId $Site.GroupId.Guid -Archived:$True -ErrorAction SilentlyContinue | Out-Null
            Write-Output "Teams for site $GroupName is set to archiving" 

        } Else {
            Write-Output "Teams for site $GroupName is already archived"
        }
        
    } Else {
        Write-Output "Teams site $GroupName does not exist"
    }
}