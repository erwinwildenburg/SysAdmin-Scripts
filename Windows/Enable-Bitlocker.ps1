[CmdletBinding()]
Param(
    [switch]$BackupToAd,
    [switch]$EnableTpm
)

function EnableTpmChip
{
    [CmdletBinding()]
    Param()
    
    $manufacturer = (Get-WmiObject -Class Win32_ComputerSystem).Manufacturer
    if ($manufacturer -like "Dell*")
    {
        Write-Verbose "Detected Dell system..."

        # Install Dell Command | Configure
        if (!(Test-Path -Path "C:\Temp\DCC.exe"))
        {
            Write-Verbose "Installing Dell Command | Configure..."
            
            Invoke-Webrequest "https://downloads.dell.com/FOLDER04457713M/1/Dell-Command-Configure_FVGF9_WIN_3.3.0.314_A00.EXE" -OutFile "C:\Temp\DCC.exe" -UseBasicParsing
            & C:\Temp\DCC.exe /s
            Remove-Item "C:\Users\Public\Desktop\Dell Command Configure Wizard.lnk"
        }
        
        # Enable TPM chip
        Write-Verbose "Setting temporary BIOS password..."
        & "C:\Program Files (x86)\Dell\Command Configure\X86_64\cctk.exe" --setuppwd=dell123

        Write-Verbose "Enabling TPM chip..."
        & "C:\Program Files (x86)\Dell\Command Configure\X86_64\cctk.exe" --tpm=on --valsetuppwd=dell123

        Write-Verbose "Activating TPM chip..."
        & "C:\Program Files (x86)\Dell\Command Configure\X86_64\cctk.exe" --tpmactivation=activate --valsetuppwd=dell123

        Write-Verbose "Removing temporary BIOS password..."
        & "C:\Program Files (x86)\Dell\Command Configure\X86_64\cctk.exe" --setuppwd= --valsetuppwd=dell123
    }
}

function EnableBitlocker
{
    [CmdletBinding()]
    Param()

    # Enable Bitlocker
    Write-Verbose "Enabling Bitlocker..."
    Enable-BitLocker -MountPoint "C:" -RecoveryPasswordProtector -SkipHardwareTest -UsedSpaceOnly
}

function BackupToAd
{
    [CmdletBinding()]
    Param()

    # Backup key to Active Directory
    Write-Verbose "Saving recovery key in Active Directory"
    $keyProtector = (Get-BitLockerVolume -MountPoint "C:").KeyProtector
    if ($keyProtector -ne $null)
    {
        Backup-BitLockerKeyProtector -MountPoint "C:" -KeyProtectorId ((Get-BitLockerVolume -MountPoint "C:").KeyProtector | Where-Object KeyProtectorType -eq RecoveryPassword).KeyProtectorId    
    }
}

if ($EnableTpm) { EnableTpmChip }
EnableBitlocker
if ($BackupToAd) { BackupToAd }
