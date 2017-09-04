# Import required modules
Import-Module ..\Shared\Connect-Office365.ps1

# Connect to Office 365
$connectedToOffice365 = Connect-Office365
if (!$connectedToOffice365) { exit }