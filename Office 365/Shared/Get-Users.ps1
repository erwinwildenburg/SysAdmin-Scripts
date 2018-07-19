Import-Module .\Shared\Get-GraphHeader.ps1

function Get-Users
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)][string] $clientId,
        [Parameter(Mandatory=$true)][string] $tenantId
    )

    # Get a list of all tenants
    $uri = "https://graph.microsoft.com/beta/$($tenantId)/users"
    $header = (Get-GraphHeader -clientId $clientId -tenantId $tenantId)
    $users = @()
    do
    {
        $temp = Invoke-RestMethod -Uri $uri -Headers $header -Method Get
        $users += $temp.value
        $uri = $temp.'@odata.nextlink'
    }
    until ($null -eq $uri)

    return $users
}