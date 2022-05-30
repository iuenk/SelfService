  <#PSScriptInfo
.VERSION 1.0
.AUTHOR Ivo Uenk
.RELEASENOTES

#>
<#
.SYNOPSIS
  New groups via PowerAutomate
.DESCRIPTION
  New groups via PowerAutomate
.NOTES
  Version:        1.0
  Author:         Ivo Uenk
  Creation Date:  2022-03-30
  Purpose/Change: New groups via PowerAutomate

  Install the following modules in Automation Accounts:
  AzureADPreview
  Microsoft.Online.SharePoint.PowerShell
  MicrosoftTeams
  ExchangeOnlineManagement

#>

# Parameters from Power Automate
Param (
  [string] $GroupName,
  [string] $GroupType,
  [string] $Department,
  #[string] $ConditionalAccessPolicy,
  [string] $AccessType,
  [string] $ExternalSharing,
  [boolean] $GuestAccess,
  [boolean] $Teams,
  [boolean] $Internal
)

# Import the necessary modules
Import-Module 'AzureADPreview'
Import-Module 'Microsoft.Online.SharePoint.PowerShell'
Import-Module 'MicrosoftTeams'
Import-Module 'ExchangeOnlineManagement'

# Global variables
$TenantID = Get-AutomationVariable -Name 'TenantId' 
$Tenant = Get-AutomationVariable -Name 'Tenant' 
$EmailDomain = Get-AutomationVariable -Name 'EmailDomain' 
$AppId = Get-AutomationVariable -Name 'AppId' 
$AppSecret = Get-AutomationVariable -Name 'AppSecret' 
$Wait=10

# Microsoft 365 variables
$Language = "en-US"
$RequireSenderAuthenticationEnabled = $true

# Teams variables
$AllowCreateUpdateRemoveTabs = $false
$AllowAddRemoveApps =$false
$AllowDeleteChannels =$false
$AllowCreateUpdateChannels =$false
$AllowCreateUpdateRemoveConnectors =$false

# Get the credential from Automation  
$credential = Get-AutomationPSCredential -Name 'AutomationCreds'  
$userName = $credential.UserName  
$securePassword = $credential.Password
$psCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userName, $securePassword

# Connect to Microsoft services
Connect-AzureAD -Credential $psCredential

if(($GroupType -eq 'Microsoft365') -or ($GroupType -eq 'Distributionlist'))
{
    Connect-ExchangeOnline -Credential $psCredential
} 
    
if($GroupType -eq 'Microsoft365')
{
    Connect-SPOService -Url "https://$($Tenant)-admin.sharepoint.com" -Credential $psCredential
}

if($Teams -eq $true)
{
    Connect-MicrosoftTeams -Credential $psCredential -ErrorAction Stop   
}

if($Department)
{
    function ConvertStaticGroupToDynamic
    {
            Param([string]$groupId, [string]$dynamicMembershipRule)
            $dynamicGroupTypeString = "DynamicMembership"
            # Existing group types
            [System.Collections.ArrayList]$groupTypes = (Get-AzureAdMsGroup -Id $groupId).GroupTypes

            if($null -ne $groupTypes -and $groupTypes.Contains($dynamicGroupTypeString))
            {
                throw "This group is already a dynamic group. Aborting conversion.";
            }
            # Add the dynamic group type to existing types
            $groupTypes.Add($dynamicGroupTypeString)

            # Modify the group properties to make static group dynamic
            Set-AzureAdMsGroup -Id $groupId -GroupTypes $groupTypes.ToArray() -MembershipRuleProcessingState "On" -MembershipRule $dynamicMembershipRule
    }
}

# Region Microsoft 365
if($GroupType -eq 'Microsoft365')
{
    $Alias = $GroupName.Replace(' ','').Replace('&','').ToLower()
    $EmailAddress = "$Alias@groups.$($EmailDomain)"
    
    # Create visitors group  
    if($GuestAccess -eq $true)
    {
        try
        { 
            New-AzureADGroup -DisplayName "sp-visitors-$Alias" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet" -Description "Read access to SharePoint site `"$GroupName`""
            Write-Output "SharePoint read only group 'sp-visitors-$Alias' is created" 
        
        } catch {
            Write-Output "Result:Failed (SharePoint read only group 'sp-visitors-$Alias' cannot be created"
            $_.Exception.Message
            $_.InvocationInfo.ScriptLineNumber
            exit 
        }
    }

    # Create Microsoft 365 group
    try
    {
        if ( $null -eq (Get-SPOSite -Limit all | Where-Object {$_.url -like "*/sites/$Alias" }) ) 
        {
            if (!($skip)) 
            {
                Write-Output "Creating UnifiedGroup '$GroupName'"
                New-UnifiedGroup -AccessType $AccessType `
                    -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled `
                    -DisplayName "$GroupName" `
                    -Language $Language `
                    -EmailAddresses "SMTP:$EmailAddress" `
                    -Alias $Alias | Out-Null
            }
            
            if($Alias.Length -gt 27)
            {
                $ShortAlias  = $Alias.Substring(0, 27)                
            } else {
                $ShortAlias =$Alias
            }

            try 
            {
                if (!($Identity = (Get-UnifiedGroup -Identity $ShortAlias`_*))) 
                {
                    $n = 1    
                    do 
                    {
                        $Identity = (Get-UnifiedGroup -Identity $ShortAlias`_*)
                        $n++
                        Write-Output "Get-UnifiedGroup -Identity $Alias : 15 seconds"
                        Start-Sleep -Seconds 15
                    } until (($Identity) -or ($n -eq 10) ) 
                }

                Set-UnifiedGroup -Identity $identity.Name -HiddenFromAddressListsEnabled $true -ErrorAction stop | Out-Null
                Set-UnifiedGroup -Identity $identity.Name -UnifiedGroupWelcomeMessageEnabled:$false -ErrorAction stop | Out-Null

                if ($Department)
                {
                    # Retrieve Azure AD extension property
                    $MyApp = (Get-AzureADApplication -SearchString "Custom Extensions").objectId
                    $E = Get-AzureADApplicationExtensionProperty -ObjectId $MyApp
                    $Extension = $E.Name.Insert(0,'user.')

                    $ID = (Get-UnifiedGroup $Identity.Name).ExternalDirectoryObjectId
                    ConvertStaticGroupToDynamic -groupId $ID -dynamicMembershipRule "(user.department -eq ""$Department"") or ($Extension -eq ""$Department"") or (user.department -eq ""Board"")"
                }

            } catch {
                Write-Error $_.Exception.Message
            }

            #enable site
            [int]$c=1
            do 
            {
                Try
                {
                    Start-Sleep -Seconds $Wait

                    $authString = "https://login.microsoftonline.com/$tenantId" 
                    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext"-ArgumentList $authString
                    $creds = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential" -ArgumentList $AppId, $AppSecret
                    $context = $authContext.AcquireTokenAsync("https://graph.microsoft.com/", $creds).Result

                    $GroupId = $Identity.ExternalDirectoryObjectId

                    Invoke-WebRequest -Headers @{
                    Authorization  = $context.CreateAuthorizationHeader()
                    'Content-Type' = "application/json"
                    } -uri https://graph.microsoft.com/v1.0/groups/$GroupId/drive/ -UseBasicParsing
                    $GroupEnabled = $true

                } catch {
                    $GroupEnabled = $false
                    $c++
                }

            } until ($GroupEnabled -or ($c -eq 10))

        } else {
            Write-Output "Site $alias allready exists, skipping creation"
        }                    
        try 
        {
            if (!($SPOIdentity = Get-SPOSite "https://$tenant.sharepoint.com/sites/$Alias")) 
            {
                $n = 1    
                do 
                {
                    $SPOIdentity = Get-SPOSite "https://$tenant.sharepoint.com/sites/$Alias" -ErrorAction SilentlyContinue
                    $n++
                    Write-Output "get-SPOSite https://$tenant.sharepoint.com/sites/$Alias  : 15 seconds"
                    Start-Sleep -Seconds 15
                } until (($SPOIdentity) -or ($n -eq 10) ) 
            }

            if ($ExternalSharing) 
            {
                Write-Output "enabling $ExternalSharing for $GroupName"
                Set-SPOSite -Identity "https://$tenant.sharepoint.com/sites/$Alias" -SharingCapability $ExternalSharing -ErrorAction Stop | Out-Null

            } else {
                Write-Output "disabling sharing capabilities for $GroupName"
                Set-SPOSite -Identity "https://$tenant.sharepoint.com/sites/$alias" -SharingCapability 'Disabled' -ErrorAction Stop | Out-Null
                Set-SPOSite -Identity "https://$tenant.sharepoint.com/sites/$alias" -DisableSharingForNonOwners:$true -ErrorAction Stop | Out-Null
            }
            #if($ConditionalAccessPolicy -eq $true){
            #    Write-Output "Setting CA Policy $ConditionalAccessPolicy for $GroupName"
            #    Set-SPOSite -Identity "https://$tenant.sharepoint.com/sites/$alias" -ConditionalAccessPolicy $ConditionalAccessPolicy -ErrorAction Stop | Out-Null
            #}

        } catch {
        Write-Output "Result:Failed (Microsoft365 group '$GroupName' not created)"
        Write-Error $_.Exception.Message
        $_.InvocationInfo.ScriptLineNumber
        exit
    }

    If ($Teams -eq $true) 
    {
        try 
        {
            Write-Output "Enabling Teams"
            New-Team -Group $Identity.ExternalDirectoryObjectId -ErrorAction Stop | out-null
            Set-Team -GroupId $Identity.ExternalDirectoryObjectId `
                -AllowCreateUpdateRemoveConnectors $AllowCreateUpdateRemoveConnectors `
                -AllowCreateUpdateChannels $AllowCreateUpdateChannels `
                -AllowDeleteChannels $AllowDeleteChannels `
                -AllowAddRemoveApps $AllowAddRemoveApps `
                -AllowCreateUpdateRemoveTabs $AllowCreateUpdateRemoveTabs `
                -ErrorAction Stop | Out-Null
                
        } catch {
            Write-Output "Result:Failed (Teams site not created)"
            $_.Exception.Message
            $_.InvocationInfo.ScriptLineNumber
            exit
        }
    }
        Write-Output "Microsoft365 group '$GroupName' is created" 
        
    } catch {
        #Do Nothing ....
    }
}

If($GuestAccess -eq $true)
{
    $SPOGroup = "$GroupName Visitors"
    $Alias = $GroupName.Replace(' ','').Replace('&','').ToLower()
    $LoginName = "sp-visitors-$Alias"
    [bool]$Success = $false
    Do
    {
        try
        {
            Add-SPOUser -Group $SPOGroup -LoginName $LoginName -Site "https://$($Tenant).sharepoint.com/sites/$Alias" -ErrorAction Stop
            Write-Output "Group '$LoginName' is added" 
            [bool]$Success = $true

        } catch {
            [bool]$Success=$false
            Write-Output "Wait 60 seconds and try again to add $LoginName"
            Start-Sleep 60
        }
    } until($Success)
}
# End region Microsoft 365

# Create Security group
if($GroupType -eq 'Security')
{
    try
    {    
        $GroupName = "sg-$($GroupName.tolower().replace(' ',''))"
        New-AzureADGroup -DisplayName $GroupName -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet" -Description "Security groep '$GroupName'" -ErrorAction Stop
        Write-Output "Security group '$GroupName' is created"

    } catch {
        Write-Output "Result:Failed (Security group '$GroupName' cannot be created)" 
    }

    if ($Department -ne $null)
    {

        # Retrieve Azure AD extension property
        $MyApp = (Get-AzureADApplication -SearchString "Custom Extensions").objectId
        $E = Get-AzureADApplicationExtensionProperty -ObjectId $MyApp
        $Extension = $E.Name.Insert(0,'user.')

        $ID = (Get-AzureADGroup -SearchString $GroupName).ObjectId
        ConvertStaticGroupToDynamic -groupId $ID -dynamicMembershipRule "(user.department -eq ""$Department"") or ($Extension -eq ""$Department"")"
    }
}

# Create Distribution group
if($GroupType -eq 'Distributionlist')
{
    try
    {
        $Alias = "dl-$($GroupName.tolower().replace(' ',''))"
        $PrimarySmtpAddress = "$Alias@$($EmailDomain)"
        New-DistributionGroup -DisplayName $GroupName -name $Alias -RequireSenderAuthenticationEnabled $Internal -PrimarySmtpAddress $PrimarySmtpAddress -ErrorAction Stop
        Write-Output "Distributionlist '$GroupName' is created" 

    } catch {
        Write-Output "Result:Failed (Distributionlist not created)"
    }
}