# Please make sure the root directory contains the correct permissions
# For more information: https://support.microsoft.com/en-us/kb/274443

# Enumerate a list of folders
$Folders = Get-ChildItem |? { $_.PSIsContainer }
foreach ($Folder in $Folders)
{
    # Recursively set owner 
    takeown.exe /F $($Folder.FullName) /R /D Y | Out-Null

    # Recursively re-enable inherited permissions on the folder
    icacls.exe $($Folder.FullName) /reset /T /C /L /Q

    # Recursively grant the user access rights to the folder
    icacls.exe $($Folder.FullName) /grant $($Folder.BaseName) + ":(OI)(CI)F"

    # Set the owner back to the user
    icacls.exe $($Folder.FullName) /setowner $($Folder.BaseName) /T /C /L /Q
}
