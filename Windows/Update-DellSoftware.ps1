# Get the device information
$manufacturer = (Get-WmiObject -Class Win32_ComputerSystem).Manufacturer
$model = (Get-WmiObject -Class Win32_ComputerSystem).Model
$is64BitOs = [System.Environment]::Is64BitOperatingSystem

# Install and run Dell Command | Update
if ($manufacturer -like "Dell*" -and (($model -like "OptiPlex*") -or ($model -like "Latitude*") -or ($model -like "Precision*") -or ($model -like "Vanue Tablets*") -or ($model -like "XPS*")))
{
    $isInstalled32 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Dell Command | Update" }
    $isInstalled64 = Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Dell Command | Update" }
    if (($isInstalled32.DisplayName -ne "Dell Command | Update") -and ($isInstalled64.DisplayName -ne "Dell Command | Update"))
    {
        Write-Host "Installing Dell Command | Update..."

        # Download Dell Command Update
        Invoke-WebRequest -Uri "https://downloads.dell.com/FOLDER04358530M/1/Dell-Command-Update_X79N4_WIN_2.3.1_A00.EXE" -OutFile "C:\Temp\Dell-Command-Update_X79N4_WIN_2.3.1_A00.EXE" -UseBasicParsing

        # Install Dell Command Update
        & "C:\Temp\Dell-Command-Update_X79N4_WIN_2.3.1_A00.EXE" /s
    }
    else
    {
        Write-Host "Dell Command | Update is already installed!"
    }

    Write-Host "Installing Dell updates..."
    if ($is64BitOs)
    {
        & "${env:ProgramFiles(x86)}\Dell\CommandUpdate\dcu-cli.exe"
    }
    else
    {
        & "${env:ProgramFiles}\Dell\CommandUpdate\dcu-cli.exe"
    }

    Write-Host "Finished installing updates from Dell!"
}
