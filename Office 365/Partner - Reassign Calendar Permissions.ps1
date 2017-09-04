# Import required modules
Import-Module ..\Shared\Connect-Office365.ps1

# Connect to Office 365
$connectedToOffice365 = Connect-Office365
if (!$connectedToOffice365) { exit }

# Get a list of all user mailboxes
$users = Get-Mailbox -RecipientType "UserMailbox"

foreach ($user in $users)
{
    # Get all calendars of the user
    $calendars = Get-MailboxFolderStatistics -Identity $user.PrimarySmtpAddress | Where-Object { $_.FolderType -eq "Calendar" }
    
    foreach ($calendar in $calendars)
    {
        Write-Host "Grabbing all permissions for calendar $($calendar.Name) of user $($user.Name)"
        $calendarIdentity = ($calendar.Identity).Replace("\", ":\")
        $calendarPermissions = Get-MailboxFolderPermission -Identity $calendarIdentity
        if ($calendarPermissions)
        {
            $calendarPermissions = $calendarPermissions | Where-Object { $_.User.ToString() -ne "Default" -and $_.User.ToString() -ne "Anonymous"}
            $calendarPermissions
            foreach ($calendarPermission in $calendarPermissions)
            {
                # Remove the user permissions from the calendar
                Remove-MailboxFolderPermission -Identity $calendarIdentity -User $calendarPermission.User.ToString() -Confirm:$false

                # Add to user permissions to the calendar again
                $userToAdd = $users | Where-Object { $_.DisplayName -eq $calendarPermission.User }
                $accessRights = $calendarPermission.AccessRights -join ","
                Add-MailboxFolderPermission -Identity $calendarIdentity -User $userToAdd.PrimarySmtpAddress -AccessRights $accessRights

                Write-Host "Fixed permissions of $($calendarPermission.User) in calendar $($calendar.Name)"
            }
        }
    }
}
