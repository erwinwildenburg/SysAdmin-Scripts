# Import requirements
Import-Module MSOnline
Add-Type -AssemblyName System.Windows.Forms

# Cleanup first
Register-EngineEvent -SourceIdentifier ([System.Management.Automation.PsEngineEvent]::Exiting) -Action {
    Get-PSSession | Remove-PSSession
    Remove-Variable -Name MSOLTenantid -Scope Global -ErrorAction SilentlyContinue
} | Out-Null
Get-PSSession | Remove-PSSession
Remove-Variable -Name MSOLTenantid -Scope Global -ErrorAction SilentlyContinue

# Connect to your own tenant
$UserCredential = Get-Credential -Credential $null
if (!$UserCredential) { exit }

# Show a "Please wait..." form in case of many partners or slow connection
$PleaseWaitForm = New-Object System.Windows.Forms.Form
$PleaseWaitForm.FormBorderStyle = 'Fixed3D'
$PleaseWaitForm.MaximizeBox = $false
$PleaseWaitForm.Width = 400
$PleaseWaitForm.Height = 100
$PleaseWaitForm.Text = "Processing..."
$PleaseWaitFormLabel = New-Object System.Windows.Forms.Label
$PleaseWaitFormLabel.Location = New-Object System.Drawing.Size(10,20) 
$PleaseWaitFormLabel.Size = New-Object System.Drawing.Size(280,20) 
$PleaseWaitFormLabel.Text = "Connecting to Office 365, please wait..."
$PleaseWaitForm.Controls.Add($PleaseWaitFormLabel)
$PleaseWaitForm.Show()
$PleaseWaitForm.BringToFront()
$PleaseWaitForm.Refresh()

# Connect to Office 365
$PleaseWaitFormLabel.Text = "Getting partner list, please wait..."
$PleaseWaitForm.Refresh()
Connect-MsolService -Credential $UserCredential -ErrorAction Stop

# Get all partner tenants
$PartnerContracts = Get-MsolPartnerContract -All

# Close the form
$PleaseWaitForm.Close()

# Form for the user to select which tenant they want to connect to
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
$PartnerListForm = New-Object System.Windows.Forms.Form
$PartnerListForm.FormBorderStyle = 'Fixed3D'
$PartnerListForm.MaximizeBox = $false
$PartnerListForm.Text = "Select a customer"
$PartnerListForm.Size = New-Object System.Drawing.Size(500,385) 
$PartnerListForm.StartPosition = "CenterScreen"
$PartnerListFormOK = New-Object System.Windows.Forms.Button
$PartnerListFormOK.Location = New-Object System.Drawing.Size(165,290)
$PartnerListFormOK.Size = New-Object System.Drawing.Size(75,23)
$PartnerListFormOK.Text = "OK"
$PartnerListFormOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$PartnerListForm.Controls.Add($PartnerListFormOK)
$PartnerListForm.AcceptButton = $PartnerListFormOK
$PartnerListFormCancel = New-Object System.Windows.Forms.Button
$PartnerListFormCancel.Location = New-Object System.Drawing.Size(240,290)
$PartnerListFormCancel.Size = New-Object System.Drawing.Size(75,23)
$PartnerListFormCancel.Text = "Cancel"
$PartnerListFormCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$PartnerListForm.Controls.Add($PartnerListFormCancel)
$PartnerListForm.CancelButton = $PartnerListFormCancel
$PartnerListFormLabel = New-Object System.Windows.Forms.Label
$PartnerListFormLabel.Location = New-Object System.Drawing.Size(10,20) 
$PartnerListFormLabel.Size = New-Object System.Drawing.Size(280,20) 
$PartnerListFormLabel.Text = "Please select a customer:"
$PartnerListForm.Controls.Add($PartnerListFormLabel) 
$PartnerListFormList = New-Object System.Windows.Forms.ListBox 
$PartnerListFormList.Location = New-Object System.Drawing.Size(10,40) 
$PartnerListFormList.Size = New-Object System.Drawing.Size(460,20) 
$PartnerListFormList.Height = 250
[void] $PartnerListFormList.Items.Add("Your own tenant")
ForEach ($Partner in ($PartnerContracts | Sort -Property Name))
{
    [void] $PartnerListFormList.Items.Add($Partner.Name)
}
$PartnerListForm.Controls.Add($PartnerListFormList)
$PartnerListFormCopyright = New-Object System.Windows.Forms.Label
$PartnerListFormCopyright.Location = New-Object System.Drawing.Size(290,320)
$PartnerListFormCopyright.Size = New-Object System.Drawing.Size(280,20)
$PartnerListFormCopyright.Text = "Copyright " + [char]0x00A9 + " 2016 Erwin Wildenburg"
$PartnerListFormCopyright.ForeColor = "Gray"
$PartnerListForm.Controls.Add($PartnerListFormCopyright)
$PartnerListForm.Topmost = $True
$PartnerListForm.Add_Shown({$PartnerListForm.Activate()})
$result = $PartnerListForm.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $PartnerListFormList.SelectedIndex -ge 0)
{
    $PartnerName = $PartnerListFormList.SelectedItem
    $PartnerTenantId = ($PartnerContracts | Where { $_.Name -eq $PartnerListFormList.SelectedItem }).TenantId.Guid

    # Connect to Office 365
    if ($PartnerListFormList.SelectedItem -eq "Your own tenant") {
        $ConnectionUri = "https://outlook.office365.com/powershell-liveid/"
    }
    else
    {
        Set-Variable -Name MSOLTenantid -Value $PartnerTenantId -Scope Global
        Write-Host "Succesfully connected to Office 365 of customer",$PartnerName -ForegroundColor Green
    }

    # Connect to Exchange Online
    $ConnectionUri = "https://ps.outlook.com/powershell-liveid?DelegatedOrg=" + $PartnerTenantId
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Credential $UserCredential -Authentication Basic -AllowRedirection -ErrorAction Stop
    if (!$Session) {
        Write-Host "Failed to connect to Exchange Online of customer",$PartnerName -ForegroundColor Red
    }
    else {
        Import-PSSession $Session -ErrorAction Stop -DisableNameChecking -AllowClobber | Out-Null
        Write-Host "Succesfully connected to Exchange Online of customer",$PartnerName -ForegroundColor Green
    }
}