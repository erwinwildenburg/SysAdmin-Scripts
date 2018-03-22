function Get-FileserverPermissions
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)][string] $path,
        [Parameter(Mandatory=$true)][string] $saveLocation
    )

    # Provide a stable path variable
    $path = Join-Path -Path $path -ChildPath ""

    # Get a list of all directories
    $childitems = Get-ChildItem -Path $path -Directory -Recurse
    
    # Get a list of all permissions for every directory
    # We need to try/catch it because Get-Acl crashes when it has no permissions on a folder
    [System.Security.AccessControl.FileSystemSecurity[]] $acls = @()
    foreach ($item in $childitems)
    {
        try {
            $acls += Get-Acl -Path $item.FullName
        }
        catch { }
    }

    # Get a list of used permissions in the filesystem
    $permissions = $acls.Access.FileSystemRights -split ", " | Sort-Object | Get-Unique

    [PSCustomObject[]] $directories = @()
    foreach ($acl in $acls)
    {
        # Get the users with access to the folder
        # We filter out any system accounts
        $access = $acl.Access | Where-Object { $_.AccessControlType -eq "Allow" -and $_.IdentityReference -match "^((?!NT AUTHORITY\\.*$)(?!BUILTIN\\)).*" }

        # Format the data in a way we can export
        $directory = [PSCustomObject] @{}
        $directory | Add-Member -MemberType NoteProperty -Name "Path" -Value (Convert-Path $acl.Path)
        foreach ($permission in $permissions)
        {
            # Get a list of users with the specific user rights on the folder
            [string] $users = (($access | Where-Object { ($_.FileSystemRights -split ", ") -contains $permission }).IdentityReference -split "; " | Sort-Object | Get-Unique) -join "; "

            # Add it to the object which contains the information we need
            $directory | Add-Member -MemberType NoteProperty -Name $permission -Value $users
        }

        # Add the object to our permission matrix
        $directories += $directory
    }

    $directories | Export-Csv -Path $saveLocation -NoTypeInformation
}

Add-Type -AssemblyName System.Windows.Forms

# Ask for the location to scan
$selectFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$selectFolderDialog.ShowDialog() | Out-Null
if ($selectFolderDialog.SelectedPath -eq "") { exit }

# Ask for the location to save the file
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveFileDialog.filter = "CSV (Comma delimited)|*.csv"
$saveFileDialog.ShowDialog() | Out-Null
if ($saveFileDialog.FileName -eq "") { exit }

# Show a waiting dialog to the user
$pleaseWaitForm = New-Object System.Windows.Forms.Form
$pleaseWaitForm.FormBorderStyle = 'Fixed3D'
$pleaseWaitForm.MaximizeBox = $false
$pleaseWaitForm.Width = 400
$pleaseWaitForm.Height = 100
$pleaseWaitForm.Text = "Processing..."
$pleaseWaitFormLabel = New-Object System.Windows.Forms.Label
$pleaseWaitFormLabel.Location = New-Object System.Drawing.Size(10,20) 
$pleaseWaitFormLabel.Size = New-Object System.Drawing.Size(280,20) 
$pleaseWaitFormLabel.Text = "Generating the permission matrix, please wait..."
$pleaseWaitForm.Controls.Add($pleaseWaitFormLabel)
$pleaseWaitForm.Show()
$pleaseWaitForm.BringToFront()
$pleaseWaitForm.Refresh()

Get-FileserverPermissions -path $selectFolderDialog.SelectedPath -saveLocation $saveFileDialog.FileName

# Close the waiting dialog
$pleaseWaitForm.Close()