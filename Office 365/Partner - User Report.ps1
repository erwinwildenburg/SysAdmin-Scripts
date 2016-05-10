# Load required assembly
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
$PleaseWaitForm.FormBorderStyle = "Fixed3D"
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
$PartnerListForm.FormBorderStyle = "Fixed3D"
$PartnerListForm.MaximizeBox = $false
$PartnerListForm.Text = "Select a customer"
$PartnerListForm.Size = New-Object System.Drawing.Size(500,385) 
$PartnerListForm.StartPosition = "CenterScreen"
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
$PartnerListFormList.SelectionMode = "MultiExtended"
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
$PartnerListFormOK = New-Object System.Windows.Forms.Button
$PartnerListFormOK.Location = New-Object System.Drawing.Size(165,290)
$PartnerListFormOK.Size = New-Object System.Drawing.Size(75,23)
$PartnerListFormOK.Text = "OK"
$PartnerListFormOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$PartnerListForm.Controls.Add($PartnerListFormOK)
$PartnerListForm.AcceptButton = $PartnerListFormOK
$PartnerListForm.Topmost = $True
$PartnerListForm.Add_Shown({$PartnerListForm.Activate()})
$result = $PartnerListForm.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $PartnerListFormList.SelectedIndex -ge 0)
{
	Foreach ($partner in $PartnerListFormList.SelectedItems)
	{
		$PartnerName = $partner
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

		# Get the information we want
		$ExportData = "UserPrincipalName,DisplayName,Office,LastLogonTime,BlockCredential,Licenses`n"
		$Office365Users = Get-MsolUser | Where-Object { $_.isLicensed -eq "TRUE" } | Select UserPrincipalName,DisplayName,Office,BlockCredential,Licenses
		Foreach ($User in $Office365Users)
		{
			# Get the user data
			$UserPrincipalName = $User.UserPrincipalName
			$LastLogonTime = (Get-MailboxStatistics -Identity $UserPrincipalName).LastLogonTime
			$DisplayName = $User.DisplayName
			$BlockCredential = $User.BlockCredential
			$Office = $User.Office
		
			# Change user data if necessary
			if ($LastLogonTime -eq $null)
			{
				$LastLogonTime = "Nooit"
			}
			else
			{
				$LastLogonTime = "{0:dd-MM-yyyy HH:mm}" -f $LastLogonTime
			}
			if ($BlockCredential -eq $False) {
				$BlockCredential = "Toegestaan"
			}
			else {
				$BlockCredential = "Geblokkeerd"
			}

			# Get the user licenses
			$SKUs = $User.Licenses.AccountSku.SkuPartNumber
			$Licenses = @()
			Foreach ($SKU in $SKUs)
			{
				Switch($SKU)
				{
					"EXCHANGE_L_STANDARD" {
						$Licenses += "Exchange Online Plan 1"
					}
					"MCOLITE" {
						$Licenses += "Lync Online Plan 1"
					}
					"SHAREPOINTLITE" {
						$Licenses += "SharePoint Online Plan 1"
					}
					"OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ" {
						$Licenses += "Office ProPlus"
					}
					"EXCHANGE_S_STANDARD_MIDMARKET" {
						$Licenses += "Exchange Online Plan 1"
					}
					"MCOSTANDARD_MIDMARKET" {
						$Licenses += "Lync Online Plan 1"
					}
					"SHAREPOINTENTERPRISE_MIDMARKET" {
						$Licenses += "SharePoint Online Plan 1"
					}
					"SHAREPOINTWAC" {
						$Licenses += "Office Online"
					}
					"OFFICESUBSCRIPTION" {
						$Licenses += "Office ProPlus"
					}
					"YAMMER_MIDSIZE" {
						$Licenses += "Yammer"
					}
					"EXCHANGE_S_STANDARD" {
						$Licenses += "Exchange Online Plan 2"
					}
					"MCOSTANDARD" {
						$Licenses += "Lync Online Plan 2"
					}
					"SHAREPOINTENTERPRISE" {
						$Licenses += "SharePoint Online Plan 2"
					}
					"SHAREPOINTWAC" {
						$Licenses += "Office Online"
					}
					"RMS_S_ENTERPRISE" {
						$Licenses += "Azure Active Directory Rights Management"
					}
					"YAMMER_ENTERPRISE" {
						$Licenses += "Yammer Enterprise"
					}
					"MCVOICECONF" {
						$Licenses += "Lync Online Plan 3"
					}
					"EXCHANGE_S_DESKLESS" {
						$Licenses += "Exchange Online Kiosk"
					}
					"SHAREPOINTDESKLESS" {
						$Licenses += "SharePoint Online Kiosk"
					}
					"YAMMER_ENTERPRISE" {
						$Licenses += "Yammer Enterprise"
					}
					"EXCHANGE_S_DESKLESS" {
						$Licenses += "Exchange Online Kiosk"
					}
					"SHAREPOINTDESKLESS" {
						$Licenses += "SharePoint Online Kiosk"
					}
					"SHAREPOINTWAC" {
						$Licenses += "Office Online"
					}
					"LITEPACK" {
						$Licenses += "Office 365 Plan P1"
					}
					"STANDARDPACK" {
						$Licenses += "Enterprise E1"
					}
					"STANDARDWOFFPACK" {
						$Licenses += "Enterprise E2"
					}
					"DESKLESSPACK" {
						$Licenses += "Office 365 Plan K1"
					}
					"EXCHANGEARCHIVE" {
						$Licenses += "Exchange Online Archiving"
					}
					"EXCHANGETELCO" {
						$Licenses += "Exchange Online POP"
					}
					"SHAREPOINTSTORAGE" {
						$Licenses += "SharePoint Online Storage"
					}
					"SHAREPOINTPARTNER" {
						$Licenses += "SharePoint Online Partner Access"
					}
					"PROJECTONLINE_PLAN_1" {
						$Licenses += "Project Online Plan 1"
					}
					"PROJECTONLINE_PLAN_2" {
						$Licenses += "Project Online Plan 2"
					}
					"PROJECT_CLIENT_SUBSCRIPTION" {
						$Licenses += "Project Pro"
					}
					"VISIO_CLIENT_SUBSCRIPTION" {
						$Licenses += "Visio Pro"
					}
					"INTUNE_A" {
						$Licenses += "Intune"
					}
					"CRMSTANDARD" {
						$Licenses += "CRM Online"
					}
					"CRMTESTINSTANCE" {
						$Licenses += "CRM Test Instance"
					}
					"ONEDRIVESTANDARD" {
						$Licenses += "OneDrive for Business"
					}
					"WACONEDRIVESTANDARD" {
						$Licenses += "OneDrive Pack"
					}
					"SQL_IS_SSIM" {
						$Licenses += "Power BI Information Services"
					}
					"BI_AZURE_P1" {
						$Licenses += "Power BI Reporting and Analytics"
					}
					"EOP_ENTERPRISE" {
						$Licenses += "Exchange Online Protection"
					}
					"PROJECT_ESSENTIALS" {
						$Licenses += "Project Lite"
					}
					"CRMIUR" {
						$Licenses += "CRM for Partners"
					}
					"NBPROFESSIONALFORCRM" {
						$Licenses += "Microsoft Social Listening Professional"
					}
					"AAD_PREMIUM" {
						$Licenses += "Azure Active Directory Premium"
					}
					"MFA_PREMIUM" {
						$Licenses += "Azure Multi-Factor Authentication"
					}
					"SMB_BUSSPREM" {
						$Licenses += "Office 365 Business Premium"
					}
					"ENTERPRISEPACK" {
						$Licenses += "Enterprise E3"
					}
					"EXCHANGESTANDARD" {
						$Licenses += "Exchange Online Plan 1"
					}
					"EXCHANGEENTERPRISE" {
						$Licenses += "Exchange Online Plan 2"
					}
					"SMB_BUSINESS_PREMIUM" {
						$Licenses += "Business Premium"
					}
					"PROJECTCLIENT" {
						$Licenses += "Project Pro"
					}
					"POWER_BI_STANDALONE" {
						$Licenses += "Microsoft Power BI"
					}
					"POWER_BI_STANDARD" {
						$Licenses += "Microsoft Power BI"
					}
					"SMB_BUSINESS_ESSENTIALS" {
						$Licenses += "Business Essentials"
					}
					"O365_BUSINESS_PREMIUM" {
						$Licenses += "Business Premium"
					}
					default {
						$Licenses += $SKU
					}
				}
			}
			$Licenses = $Licenses -join "; "

			$ExportData += "$UserPrincipalName,$DisplayName,$Office,$LastLogonTime,$BlockCredential,$Licenses`n"
		}

		# TEMP: Show data on console
		$ExportData = $ExportData | ConvertFrom-CSV
		$SaveFile = "C:\Temp\report.xlsx"

		# Convert data to Excel
		$Excel = New-Object -ComObject Excel.Application 
		$Excel.visible = $false
		$Excel.DisplayAlerts = $false
		$Workbooks = $Excel.Workbooks.Add()
		$Worksheets = $Workbooks.worksheets
		$Worksheet = $Worksheets.Item(1)
		$Worksheet.Name = "Office 365"

		# Add headers
		$Worksheet.Cells.Item(1,1) = "Gebruikersnaam"
		$Worksheet.Cells.Item(1,2) = "Weergave naam"
		$Worksheet.Cells.Item(1,3) = "Kantoor"
		$Worksheet.Cells.Item(1,4) = "Laatste inlog"
		$Worksheet.Cells.Item(1,5) = "Aanmeldstatus"
		$Worksheet.Cells.Item(1,6) = "Licenties"
		$Worksheet.Cells.Item(1,7) = "Verwijderen zonder archiveren"
		$Worksheet.Cells.Item(1,8) = "Toegang blokkeren"
		$Worksheet.Cells.Item(1,9) = "Toegang blokkeren en licentie aanpassen"
		$Worksheet.Cells.Item(1,10) = "Geen actie"

		# Add data
		$i = 2
		Foreach($Row in $ExportData)
		{
			$Worksheet.Cells.Item($i,1) = $Row.UserPrincipalName
			$Worksheet.Cells.Item($i,2) = $Row.DisplayName
			$Worksheet.Cells.Item($i,3) = $Row.Office
			$Worksheet.Cells.Item($i,4) = $Row.LastLogonTime
			$Worksheet.Cells.Item($i,5) = $Row.BlockCredential
			$Worksheet.Cells.Item($i,6) = $Row.Licenses
			$i++
		}

		# Autofit all data
		$Worksheet.columns.item("d").NumberFormat = "dd-mm-jjjj uu:mm"
		$Worksheet.UsedRange.EntireColumn.AutoFit()

		# Format data as table
		$ExcelTable = $Worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $Worksheet.UsedRange, 0, 1)
		$ExcelTable.Name = "Table1"
		$ExcelTable.TableStyle = "TableStyleMedium2"

		# Prompt to save file
		$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
		$SaveFileDialog.filter = "Excel-werkmap|*.xlsx"
		$SaveFileDialog.InitialDirectory = "X:\"
		$SaveFileDialog.FileName = "$PartnerName - Gebruikers accounts Office 365.xlsx"
		$SaveFileDialog.ShowDialog() | Out-Null
		If ($SaveFileDialog.FileName -ne "") {
			$Workbooks.saveas($SaveFileDialog.FileName)
		}

		$Excel.Quit()
		Remove-Variable -Name Excel
		[gc]::collect() 
		[gc]::WaitForPendingFinalizers()

		# Cleanup the session to Exchange Online
		Get-PSSession | Remove-PSSession
	}
}