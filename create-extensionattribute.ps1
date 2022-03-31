# Create a new extension property
$MyApp = (New-AzureADApplication -DisplayName "Custom Extensions" -IdentifierUris "https://ucorp.nl").ObjectId
New-AzureADServicePrincipal -AppId (Get-AzureADApplication -SearchString "Custom Extensions").AppId

# New extension property
$MyApp = (Get-AzureADApplication -SearchString "Custom Extensions").objectId
New-AzureADApplicationExtensionProperty -ObjectId $MyApp -Name "OtherDepartment" -DataType "String" -TargetObjects "User"

# Retrieve extension property
$MyApp = (Get-AzureADApplication -SearchString "Custom Extensions").objectId
$E = Get-AzureADApplicationExtensionProperty -ObjectId $MyApp
$Extension = $E.Name
 
# Set extension attribute for user
# in app dropdown maken
$UserId = "<User UPN>"
$OtherDepartment = "Developers"

Set-AzureADUserExtension -ObjectId $UserId -ExtensionName $Extension -ExtensionValue $OtherDepartment

# Check if set correctly
(Get-AzureADUserExtension -ObjectId $UserId).get_item($Extension)
# If result not empty, success