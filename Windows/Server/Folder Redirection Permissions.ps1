# Please make sure the root directory contains the correct permissions
# For more information: https://support.microsoft.com/en-us/kb/274443

# Enumerate a list of folders
$folders = Get-ChildItem | Where-Object { $_.PSIsContainer }
foreach ($folder in $folders)
{
    # Recursively set owner 
    Invoke-Command "takeown.exe /F $($folder.FullName) /R /D Y" | Out-Null

    # Recursively re-enable inherited permissions on the folder
    Invoke-Command "icacls.exe $($folder.FullName) /reset /T /C /L /Q"

    # Recursively grant the user access rights to the folder
    Invoke-Command "icacls.exe $($folder.FullName) /grant $($folder.BaseName):(OI)(CI)F /C /L /Q"

    # Set the owner back to the user
    Invoke-Command "icacls.exe $($folder.FullName) /setowner $($folder.BaseName) /T /C /L /Q"
}
