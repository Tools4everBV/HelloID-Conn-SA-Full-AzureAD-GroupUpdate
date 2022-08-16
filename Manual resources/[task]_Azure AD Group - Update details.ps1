# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

#Mapping variables from form
$description = $form.description;
$displayName = $form.displayName;
$groupId = $form.gridGroups.id;

#Change mapping here
$group = [PSCustomObject]@{
    description = $description;
    displayName = $displayName;
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
    Write-Information "Generating Microsoft Graph API Access Token user.."

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
         
    Write-Information "Searching for AzureAD group ID=$groupId"

    #Add the authorization header to the request
    $authorization = @{
        Authorization = "Bearer $accesstoken";
        'Content-Type' = "application/json";
        Accept = "application/json";
    }

    $baseSearchUri = "https://graph.microsoft.com/"
    $searchUri = $baseSearchUri + "v1.0/groups/$groupId"
    $azureADGroup = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false

    Write-Information "Finished searching AzureAD group [$groupId]"
    Write-Information "Updating AzureAD group [$($azureADGroup.displayName)].."

    $baseUpdateUri = "https://graph.microsoft.com/"
    $updateUri = $baseUpdateUri + "v1.0/groups/$($azureADGroup.id)"
    $body = $group | ConvertTo-Json -Depth 10
 
    $response = Invoke-RestMethod -Uri $updateUri -Method PATCH -Headers $authorization -Body $body -Verbose:$false

    Write-Information "AzureAD group [$($azureADGroup.displayName)] updated successfully"
    
    $Log = @{
        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
        System            = "AzureActiveDirectory" # optional (free format text) 
        Message           = "Updated group with id $groupId" # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $($azureADGroup.displayName) # optional (free format text) 
        TargetIdentifier  = $groupId # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log

}catch{
    Write-Error "Error updating AzureAD group [$($azureADGroup.displayName)]. Error: $_"     

    $Log = @{
        Action            = "UpdateResource" # optional. ENUM (undefined = default) 
        System            = "AzureActiveDirectory" # optional (free format text) 
        Message           = "Updated group with id $groupId" # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $($azureADGroup.displayName) # optional (free format text) 
        TargetIdentifier  = $groupId # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
