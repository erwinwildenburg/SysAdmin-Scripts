Import-Module .\Shared\Get-GraphHeader.ps1

function Get-PartnerContracts
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)][string] $clientId
    )

    # Get a list of all tenants
    $uri = "https://graph.microsoft.com/beta/contracts"
    $partners = @()
    do
    {
        $temp = Invoke-RestMethod -Uri $uri -Headers (Get-GraphHeader -clientId $clientId) -Method Get
        $partners += $temp.value
        $uri = $temp.'@odata.nextlink'
    }
    until ($null -eq $uri)

    return $partners
}