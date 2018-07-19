function Get-GraphHeader
{
    [CmdletBinding()]
    Param(
        [string] $clientId = "adf8d16c-1677-44e9-a1e7-5d159ca19b05",
        [System.Management.Automation.PSCredential] $credentials,
        [string] $tenantId = "common"
    )

    Import-Module Azure

    $resourceAppIdUri = "https://graph.microsoft.com"
    $authority = "https://login.microsoftonline.com/$($tenantId)"
    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"

    # Acquire the token
    try
    {
        $token = $authContext.AcquireToken($resourceAppIdUri, $clientId, $redirectUri, [Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Auto)
    }
    catch
    {
        $token = $authContext.AcquireToken($resourceAppIdUri, $clientId, $redirectUri, [Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Always)
    }

    return @{
        'Content'='application/json'
        'Authorization'=$token.CreateAuthorizationHeader()
    }
}