<#PSScriptInfo
.VERSION 1.0
.AUTHOR Ivo Uenk
.RELEASENOTES

#>
<#
.SYNOPSIS
  Device Management via PowerAutomate
.DESCRIPTION
  Device Management via PowerAutomate
.NOTES
  Version:        1.0
  Author:         Ivo Uenk
  Creation Date:  2022-03-30
  Purpose/Change: Device Management via PowerAutomate

  Install the following modules:
  AzureADPreview

#>

# Parameters from Power Automate
Param (
  [string] $UserPrincipalName,
  [string] $UserDevice,
  [string] $ChangeDeviceUser,
  [string] $Action,
  [boolean] $WipeBYOD
)

# Import the necessary modules
Import-Module 'AzureADPreview'

# Global variables
$TenantID = Get-AutomationVariable -Name 'TenantId' 
$AppId = Get-AutomationVariable -Name 'AppId' 
$AppSecret = Get-AutomationVariable -Name 'AppSecret'

# Get the credential from Automation  
$credential = Get-AutomationPSCredential -Name 'AutomationCreds'  
$userName = $credential.UserName  
$securePassword = $credential.Password
$psCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userName, $securePassword

# Connect to Microsoft services
Connect-AzureAD -Credential $psCredential

# Get MS graph API connection
$authString = "https://login.microsoftonline.com/$tenantId" 
$authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext"-ArgumentList $authString
$creds = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential" -ArgumentList $AppId, $AppSecret
$context = $authContext.AcquireTokenAsync("https://graph.microsoft.com/", $creds).Result
$AccessToken = $context.AccessToken

function Get-ManagedDevice {

[cmdletbinding()]

param
(

    $ChangeDeviceUser
)

    try {
        if($ChangeDeviceUser){
            $Resource = "deviceManagement/managedDevices?`$filter=deviceName eq '$ChangeDeviceUser'"
            $uri = "https://graph.microsoft.com/beta/$($Resource)" 

            (Invoke-RestMethod -Uri $uri -Headers @{"Authorization" = "Bearer $AccessToken"} -Method Get).value
        }
        else {
            $Resource = "deviceManagement/managedDevices?`$filter=(((deviceType%20eq%20%27desktop%27)%20or%20(deviceType%20eq%20%27windowsRT%27)%20or%20(deviceType%20eq%20%27winEmbedded%27)%20or%20(deviceType%20eq%20%27surfaceHub%27)))"
            $uri = "https://graph.microsoft.com/beta/$($Resource)"
        
            (Invoke-RestMethod -Uri $uri -Headers @{"Authorization" = "Bearer $AccessToken"} -Method Get).value
        }
    } catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Output "Response content:`n$responseBody"
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        throw "Get-ManagedDevices error"
    }

}

function Get-DevicePrimaryUser {

[cmdletbinding()]

param
(
    $deviceId
)

    $Resource = "deviceManagement/managedDevices"
    $uri = "https://graph.microsoft.com/beta/$($Resource)" + "/" + $deviceId + "/users"

    try {
        $primaryUser = Invoke-RestMethod -Uri $uri -Headers @{"Authorization" = "Bearer $AccessToken"} -Method Get
        return $primaryUser.value."id"
        
    } catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Output "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        throw "Get-DevicePrimaryUser error"
    }
}

function Set-DevicePrimaryUser {
    
[cmdletbinding()]

param
(
    $IntuneDeviceId,
    $userId
)

    $Resource = "deviceManagement/managedDevices('$IntuneDeviceId')/users/`$ref"

    try {
        $uri = "https://graph.microsoft.com/beta/$($Resource)"
        $userUri = "https://graph.microsoft.com/beta/users/" + $userId
        $id = "@odata.id"
        $JSON = @{ $id="$userUri" } | ConvertTo-Json -Compress

        Invoke-RestMethod -Uri $uri -Headers @{"Authorization" = "Bearer $AccessToken"} -Method Post -Body $JSON -ContentType "application/json"

    } catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Output "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        throw "Set-DevicePrimaryUser error"
    }   
}

function Delete-DevicePrimaryUser {

[cmdletbinding()]

param
(
    $DeviceId
)

    $Resource = "deviceManagement/managedDevices('$DeviceId')/users/`$ref"

    try {
        $uri = "https://graph.microsoft.com/beta/$($Resource)"
        Invoke-RestMethod -Uri $uri -Headers @{"Authorization" = "Bearer $AccessToken"} -Method Delete
    }

    catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        throw "Delete-DevicePrimaryUser error"
    }
}

Function Get-AADUserDevices(){

[cmdletbinding()]

param
(
    $UserID
)

$Resource = "users/$UserID/managedDevices"

    try {
    $uri = "https://graph.microsoft.com/beta/$($Resource)"
    Write-Verbose $uri
    (Invoke-RestMethod -Uri $uri -Headers @{"Authorization" = "Bearer $AccessToken"} -Method Get).Value
    }

    catch {
    $ex = $_.Exception
    $errorResponse = $ex.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();
    Write-Host "Response content:`n$responseBody" -f Red
    Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    break
    }
}

Function Wipe-BYODDevices(){

[cmdletbinding()]

param
(
    $UserID
)

$Resource = "users/$UserID/managedAppRegistrations"

    try {
    $uri = "https://graph.microsoft.com/beta/$($Resource)"
    Write-Verbose $uri
    $ManagedAppReg = (Invoke-RestMethod -Uri $uri -Headers @{"Authorization" = "Bearer $AccessToken"} -Method Get).Value
    }

    catch {
    $ex = $_.Exception
    $errorResponse = $ex.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();
    Write-Host "Response content:`n$responseBody" -f Red
    Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    break
    }

    if($ManagedAppReg){

    Write-Output "$($ManagedAppReg.Count) Managed applications found for $UserPrincipalName"
    $DeviceTag = $ManagedAppReg.deviceTag | Sort-Object -Unique
    $DeviceName = $ManagedAppReg.deviceName
        
$JSON = @"

    {
        "deviceTag": "$DeviceTag"
    }

"@
    
        # Will do selective wipe
        try {
        $uri = "https://graph.microsoft.com/beta/users/$userId/wipeManagedAppRegistrationByDeviceTag"            
        (Invoke-RestMethod -Uri $uri -Headers -Body $JSON @{"Authorization" = "Bearer $AccessToken"} -Method Post).Value
        Write-Output "wipe application data on device $DeviceName"
        }

        catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        break
        }
    }
}

Function Invoke-DeviceAction(){

[cmdletbinding()]

param
(
    [switch]$Wipe,
    [switch]$Retire,
    [switch]$Delete,
    [switch]$Sync,
    [string]$DeviceID
)

    try {
        if($Wipe){
        $Resource = "deviceManagement/managedDevices/$DeviceID/wipe"
        $uri = "https://graph.microsoft.com/beta/$($resource)"
        write-verbose $uri
        Write-Verbose "Sending wipe command to $DeviceID"
        Invoke-RestMethod -Uri $uri -Headers @{"Authorization" = "Bearer $AccessToken"} -Method Post
        }

        if($Retire){
        $Resource = "deviceManagement/managedDevices/$DeviceID/retire"
        $uri = "https://graph.microsoft.com/beta/$($resource)"
        write-verbose $uri
        Write-Verbose "Sending retire command to $DeviceID"
        Invoke-RestMethod -Uri $uri -Headers @{"Authorization" = "Bearer $AccessToken"} -Method Post
        }

        if($Delete){
        $Resource = "deviceManagement/managedDevices('$DeviceID')"
        $uri = "https://graph.microsoft.com/beta/$($resource)"
        write-verbose $uri
        Write-Verbose "Sending delete command to $DeviceID"
        Invoke-RestMethod -Uri $uri -Headers @{"Authorization" = "Bearer $AccessToken"} -Method Delete
        }
        
        if($Sync){
        $Resource = "deviceManagement/managedDevices('$DeviceID')/syncDevice"
        $uri = "https://graph.microsoft.com/beta/$($resource)"
        write-verbose $uri
        Write-Verbose "Sending sync command to $DeviceID"
        Invoke-RestMethod -Uri $uri -Headers @{"Authorization" = "Bearer $AccessToken"} -Method Post
        }
    }

    catch {
    $ex = $_.Exception
    $errorResponse = $ex.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();
    Write-Output "Response content:`n$responseBody" -f Red
    Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    break
    }
}

Function Get-AADUser(){

[cmdletbinding()]

param
(
    $userPrincipalName,
    $Property
)

$User_resource = "users"
    
    try {
        
        if($userPrincipalName -eq "" -or $null -eq $userPrincipalName){     
        $uri = "https://graph.microsoft.com/beta/$($User_resource)"
        (Invoke-RestMethod -Uri $uri -Headers @{"Authorization" = "Bearer $AccessToken"} -Method Get).Value    
        }

        else {         
            if($Property -eq "" -or $Property -eq $null){
            $uri = "https://graph.microsoft.com/beta/$($User_resource)/$userPrincipalName"
            Write-Verbose $uri
            Invoke-RestMethod -Uri $uri -Headers @{"Authorization" = "Bearer $AccessToken"} -Method Get
            }

            else {
            $uri = "https://graph.microsoft.com/beta/$($User_resource)/$userPrincipalName/$Property"
            Write-Verbose $uri
            (Invoke-RestMethod -Uri $uri -Headers @{"Authorization" = "Bearer $AccessToken"} -Method Get).Value
            }
        }
    }

    catch {
    $ex = $_.Exception
    $errorResponse = $ex.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();
    Write-Output "Response content:`n$responseBody"
    Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    break
    }
}

# Change managed device primary user
if($ChangeDeviceUser)
{
    $Device = Get-ManagedDevice -ChangeDeviceUser "$ChangeDeviceUser"

    # Set primary device user
    if($UserPrincipalName)
    {
        Write-Output "Device name:" $device."deviceName"
        $DevicePrimaryUser = Get-DevicePrimaryUser -deviceId $Device.id

        if($null -eq $DevicePrimaryUser)
        {
            Write-Output "No Intune Primary User Id set for Intune Managed Device" $Device."deviceName"
        } else {
            Write-Output "Intune Device Primary User:" $DevicePrimaryUser
        }

        $User = Get-AADUser -userPrincipalName $UserPrincipalName
        $AADUserName = $User.displayName
    
        if($DevicePrimaryUser -notmatch $User.id)
        {
            $SetDevicePrimaryUser = Set-DevicePrimaryUser -IntuneDeviceId $Device.id -userId $User.id

            if($SetDevicePrimaryUser -eq "")
            {
                Write-Output "User"$User.displayName"set as Primary User for device '$ChangeDeviceUser'..."
            }
        } else {
            Write-Output "The user '$AADUserName' specified is already the Primary User on the device..."
        }

    } else {
        # Delete primary device user
        Write-Output "No user specified"
        Write-Output "Remove primary user for device '$ChangeDeviceUser'..."

        $DevicePrimaryUser = Get-DevicePrimaryUser -deviceId $Device.id
        Write-Output "Device Primary User: $($Device.userDisplayName)"

        $DeleteDevicePrimaryUser = Delete-DevicePrimaryUser -DeviceId $Device.id

        if($DeleteDevicePrimaryUser -eq "")
        {
            Write-Output "User $($Device.userDisplayName) deleted as Primary User from the device '$ChangeDeviceUser'..."
        }
    }
}

# Invoke managed device action
if((($Action -ne $null) -and ($UserPrincipalName -ne $null) -and ($UserDevice -ne $null)))
{
    # Get user ID
    $User = Get-AADUser -userPrincipalName $UserPrincipalName
    $id = $User.Id

    # Get user Device info
    $Devices = Get-AADUserDevices -UserID $id

    $SelectedDevice = $Devices | Where-Object { $_.deviceName -eq "$UserDevice" }
    $SelectedDeviceId = $SelectedDevice | Select-Object -ExpandProperty id

    # Invoke Action
    Write-Output "Invoke action '$Action' on device '$UserDevice'..."

    if($Action -eq "Sync")
    {
        Invoke-DeviceAction -DeviceID $SelectedDeviceId -Sync -Verbose
    }

    if($Action -eq "Delete")
    {
        Invoke-DeviceAction -DeviceID $SelectedDeviceId -Delete -Verbose
    }

    if($Action -eq "Wipe")
    {
        Invoke-DeviceAction -DeviceID $SelectedDeviceId -Wipe -Verbose
    }

    if($Action -eq "Retire")
    {
        Invoke-DeviceAction -DeviceID $SelectedDeviceId -Retire -Verbose
    }

} Else {
    Write-Output "Not all necessary parameters supplied"
}

if(($WipeBYOD -eq $true) -and ($UserPrincipalName -ne $null))
{
    # Get user ID
    $User = Get-AADUser -userPrincipalName $UserPrincipalName
    $id = $User.Id

    # Selective wipe all devices
    Wipe-BYODDevices -UserID $id
}


