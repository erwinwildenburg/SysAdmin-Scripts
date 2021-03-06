# Import required modules
Import-Module ..\Shared\Connect-Azure.ps1

# Ask if the user wants to connect to Exchange Online
while ($connectToExchange -ne $true -and $connectToExchange -ne $false ) { 
    Write-Host $connectToExchange
    $temp = Read-Host -Prompt "Connect to Exchange Online? [y/n]"
    if ($temp -notmatch "[yYnN]") { continue }
    elseif ($temp -match "[yY]") { $connectToExchange = $true }
    else { $connectToExchange = $false }
}

# Connect to Office 365
$connectedToOffice365 = Connect-Azure -connectToExchange $connectToExchange
if (!$connectedToOffice365) { exit }