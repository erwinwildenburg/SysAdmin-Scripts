function Connect-Office365
{
    [CmdletBinding()]
    param()

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
    $userCredential = Get-Credential -Credential $null
    if (!$userCredential) { return $false }

    # Show a "Please wait..." form in case of many partners or slow connection
    $pleaseWaitForm = New-Object System.Windows.Forms.Form
    $pleaseWaitForm.FormBorderStyle = 'Fixed3D'
    $pleaseWaitForm.MaximizeBox = $false
    $pleaseWaitForm.Width = 400
    $pleaseWaitForm.Height = 100
    $pleaseWaitForm.Text = "Processing..."
    $pleaseWaitFormLabel = New-Object System.Windows.Forms.Label
    $pleaseWaitFormLabel.Location = New-Object System.Drawing.Size(10,20) 
    $pleaseWaitFormLabel.Size = New-Object System.Drawing.Size(280,20) 
    $pleaseWaitFormLabel.Text = "Connecting to Office 365, please wait..."
    $pleaseWaitForm.Controls.Add($pleaseWaitFormLabel)
    $pleaseWaitForm.Show()
    $pleaseWaitForm.BringToFront()
    $pleaseWaitForm.Refresh()

    # Connect to Office 365
    $pleaseWaitFormLabel.Text = "Getting partner list, please wait..."
    $pleaseWaitForm.Refresh()
    Connect-MsolService -Credential $userCredential -ErrorAction Stop

    # Get all partner tenants
    $partnerContracts = Get-MsolPartnerContract -All

    # Close the form
    $pleaseWaitForm.Close()

    # Form for the user to select which tenant they want to connect to
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    $partnerListForm = New-Object System.Windows.Forms.Form
    $partnerListForm.FormBorderStyle = 'Fixed3D'
    $partnerListForm.MaximizeBox = $false
    $partnerListForm.Text = "Select a customer"
    $partnerListForm.Size = New-Object System.Drawing.Size(500,385) 
    $partnerListForm.StartPosition = "CenterScreen"
    $partnerListFormOK = New-Object System.Windows.Forms.Button
    $partnerListFormOK.Location = New-Object System.Drawing.Size(165,290)
    $partnerListFormOK.Size = New-Object System.Drawing.Size(75,23)
    $partnerListFormOK.Text = "OK"
    $partnerListFormOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $partnerListForm.Controls.Add($partnerListFormOK)
    $partnerListForm.AcceptButton = $partnerListFormOK
    $partnerListFormCancel = New-Object System.Windows.Forms.Button
    $partnerListFormCancel.Location = New-Object System.Drawing.Size(240,290)
    $partnerListFormCancel.Size = New-Object System.Drawing.Size(75,23)
    $partnerListFormCancel.Text = "Cancel"
    $partnerListFormCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $partnerListForm.Controls.Add($partnerListFormCancel)
    $partnerListForm.CancelButton = $partnerListFormCancel
    $partnerListFormLabel = New-Object System.Windows.Forms.Label
    $partnerListFormLabel.Location = New-Object System.Drawing.Size(10,20) 
    $partnerListFormLabel.Size = New-Object System.Drawing.Size(280,20) 
    $partnerListFormLabel.Text = "Please select a customer:"
    $partnerListForm.Controls.Add($partnerListFormLabel) 
    $partnerListFormList = New-Object System.Windows.Forms.ListBox 
    $partnerListFormList.Location = New-Object System.Drawing.Size(10,40) 
    $partnerListFormList.SelectionMode = "MultiExtended"
    $partnerListFormList.Size = New-Object System.Drawing.Size(460,20) 
    $partnerListFormList.Height = 250
    [void] $partnerListFormList.Items.Add("Your own tenant")
    foreach ($partner in ($partnerContracts | Sort-Object -Property Name))
    {
        [void] $partnerListFormList.Items.Add($partner.Name)
    }
    $partnerListForm.Controls.Add($partnerListFormList)
    $partnerListFormCopyright = New-Object System.Windows.Forms.Label
    $partnerListFormCopyright.Location = New-Object System.Drawing.Size(290,320)
    $partnerListFormCopyright.Size = New-Object System.Drawing.Size(280,20)
    $partnerListFormCopyright.Text = "Copyright " + [char]0x00A9 + " 2017 Erwin Wildenburg"
    $partnerListFormCopyright.ForeColor = "Gray"
    $partnerListForm.Controls.Add($partnerListFormCopyright)
    $partnerListForm.Topmost = $True
    $partnerListForm.Add_Shown({$partnerListForm.Activate()})
    $result = $partnerListForm.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $partnerListFormList.SelectedIndex -ge 0)
    {
        foreach ($partnerName in $partnerListFormList.SelectedItems)
        {
            $partnerTenantId = ($partnerContracts | Where-Object { $_.Name -eq $partnerName }).TenantId.Guid

            # Connect to Office 365
            if ($partnerName -eq "Your own tenant") {
                $connectionUri = "https://outlook.office365.com/powershell-liveid/"
            }
            else
            {
                Set-Variable -Name MSOLTenantid -Value $partnerTenantId -Scope Global
                Write-Host "Succesfully connected to Office 365 of customer",$partnerName -ForegroundColor Green
            }

            # Connect to Exchange Online
            $connectionUri = "https://ps.outlook.com/powershell-liveid?DelegatedOrg=" + $partnerTenantId
            $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Credential $userCredential -Authentication Basic -AllowRedirection -ErrorAction Stop
            if (!$session) {
                Write-Host "Failed to connect to Exchange Online of customer",$partnerName -ForegroundColor Red
                return $false
            }
            else {
                Import-PSSession $session -ErrorAction Stop -DisableNameChecking -AllowClobber | Out-Null
                Write-Host "Succesfully connected to Exchange Online of customer",$partnerName -ForegroundColor Green
                return $true
            }
        }
    }
}