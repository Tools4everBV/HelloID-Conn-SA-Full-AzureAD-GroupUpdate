#$groupTypes = @("")

#$allowExternalSenders = $false
#$autoSubscribeNewMembers = $false

#$visibility = "Private"

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

#Change mapping here
$group = [PSCustomObject]@{
    description = $description;
    displayName = $displayName;

    #groupTypes = @($groupTypes);

    mailEnabled = $mailEnabled;
    mailNickname = $mailNickName;
    # allowExternalSenders = $allowExternalSenders; - Not supported with Application permissions
    # autoSubscribeNewMembers = $autoSubscribeNewMembers; - Not supported with Application permissions

    securityEnabled = $securityEnabled;

    #visibility = $visibility;
}

# Filter out empty properties
$groupTemp = $group

$group = @{}
foreach($property in $groupTemp.PSObject.Properties){
    if(![string]::IsNullOrEmpty($property.Value)){
        $null = $group.Add($property.Name, $property.Value)
    }
}

$group = [PSCustomObject]$group

try{
    Hid-Write-Status -Message "Generating Microsoft Graph API Access Token user.." -Event Information

    $baseUri = "https://login.microsoftonline.com/"
    $authUri = $baseUri + "$AADTenantID/oauth2/token"

    $body = @{
        grant_type      = "client_credentials"
        client_id       = "$AADAppId"
        client_secret   = "$AADAppSecret"
        resource        = "https://graph.microsoft.com"
    }
 
    $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
    $accessToken = $Response.access_token;
         
    Hid-Write-Status -Message "Searching for AzureAD group ID=$groupId" -Event Information

    #Add the authorization header to the request
    $authorization = @{
        Authorization = "Bearer $accesstoken";
        'Content-Type' = "application/json";
        Accept = "application/json";
    }

    $baseSearchUri = "https://graph.microsoft.com/"
    $searchUri = $baseSearchUri + "v1.0/groups/$groupId"
    $azureADGroup = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
    HID-Write-Status -Message "Finished searching AzureAD group [$groupId]" -Event Information

    Hid-Write-Status -Message "Updating AzureAD group [$($azureADGroup.displayName)].." -Event Information

    $baseUpdateUri = "https://graph.microsoft.com/"
    $updateUri = $baseUpdateUri + "v1.0/groups/$($azureADGroup.id)"
    $body = $group | ConvertTo-Json -Depth 10
 
    $response = Invoke-RestMethod -Uri $updateUri -Method PATCH -Headers $authorization -Body $body -Verbose:$false

    Hid-Write-Status -Message "AzureAD group [$($azureADGroup.displayName)] updated successfully" -Event Success
    HID-Write-Summary -Message "AzureAD group [$($azureADGroup.displayName)] updated successfully" -Event Success
}catch{
    HID-Write-Status -Message "Error updating AzureAD group [$($azureADGroup.displayName)]. Error: $_" -Event Error
    HID-Write-Summary -Message "Error updating AzureAD group [$($azureADGroup.displayName)]" -Event Failed
}
