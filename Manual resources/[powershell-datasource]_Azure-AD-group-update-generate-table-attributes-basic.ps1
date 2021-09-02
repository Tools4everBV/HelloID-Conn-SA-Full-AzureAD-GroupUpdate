# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

try {
    $groupId = $datasource.selectedGroup.id

    Write-Information "Generating Microsoft Graph API Access Token.."

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
         
    Write-Information "Searching for AzureAD group id=$groupId"

    #Add the authorization header to the request
    $authorization = @{
        Authorization = "Bearer $accesstoken";
        'Content-Type' = "application/json";
        Accept = "application/json";
    }
 
    $baseSearchUri = "https://graph.microsoft.com/"
    $searchUri = $baseSearchUri + "v1.0/groups/$groupId"
    $azureADGroup = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
    Write-Information "Finished searching AzureAD group [$($azureADGroup.displayName)]"

    # Properties to display
    $properties = @("displayName","description","groupTypes","mailEnabled","mailNickName","allowExternalSenders","autoSubscribeNewMembers","securityEnabled","visibility")
    
    # Get first entry of groupTypes (HelloID does not support multivalue)
    $azureADGroup.groupTypes = $azureADGroup.groupTypes[0]

    foreach($tmp in $azureADGroup.psObject.properties)
    {
        if($tmp.Name -in $properties){
            $returnObject = [Ordered]@{
                name=$tmp.Name;
                value=$tmp.value
            }
            
            Write-Output $returnObject
        }
    }

    Write-Information "Finished retrieving AzureAD group [$($azureADGroup.displayName)] basic attributes" 
} catch {
    Write-Error "Error searching for AzureAD group [$groupId]. Error: $($_.Exception.Message)"
}
