# Import required modules
Import-Module ..\Shared\Connect-Azure.ps1

# Connect to Office 365
$connectedToOffice365 = Connect-Azure -connectToExchange $true
if (!$connectedToOffice365) { exit }

# Get the information we want
$exportData = "UserPrincipalName,DisplayName,Office,LastLogonTime,AccountEnabled,PasswordNeverExpires,Licenses`n"
$office365Users = Get-AzureADUser -All $true | Select-Object UserPrincipalName,DisplayName,PhysicalDeliveryOfficeName,AccountEnabled,PasswordNeverExpire,AssignedLicenses
foreach ($user in $office365Users)
{
	# Get the user data
	$userPrincipalName = $user.UserPrincipalName
	$lastLogonTime = (Get-MailboxStatistics -Identity $userPrincipalName -ErrorAction SilentlyContinue).LastLogonTime
	$displayName = $user.DisplayName
	$accountEnabled = $user.AccountEnabled
	$office = $user.PhysicalDeliveryOfficeName

	# Change user data if necessary
	if ($lastLogonTime -eq $null)
	{
		$lastLogonTime = "Never"
	}
	else
	{
		$lastLogonTime = [DateTime]::ParseExact($lastLogonTime, "MM/dd/yyyy HH:mm:ss", $null).ToString("dd-MM-yyyy HH:mm")
	}
	if ($accountEnabled -eq $true) {
		$accountEnabled = "Allowed"
	}
	else {
		$accountEnabled = "Blocked"
	}
	if ($user.PasswordNeverExpires -contains "DisablePasswordExpiration")
	{
		$passwordExpiration = "No"
	}
	else {
		$passwordExpiration = "Yes"
	}

	# Get the user licenses
	$script:skuPartNames = @()
	$skuParts = Get-AzureADSubscribedSku | Select-Object SkuId,SkuPartNumber
	$user.AssignedLicenses.SkuId | ForEach-Object {
		$partNumber = $_
		$script:skuPartNames += ($skuParts | Where-Object { $_.SkuId -eq $partNumber }).SkuPartNumber
	}

	# Translate the user licenses
	$licenses = @()
	foreach ($sku in $skuPartNames)
	{
		switch($sku)
		{
			"EXCHANGE_L_STANDARD" { $licenses += "Exchange Online (Plan 1)" }
			"MCOLITE" { $licenses += "Lync Online (Plan 1)" }
			"SHAREPOINTLITE" { $licenses += "SharePoint Online (Plan 1)" }
			"OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ" { $licenses += "Office Pro Plus" }
			"EXCHANGE_S_STANDARD_MIDMARKET" { $licenses += "Exchange Online (Plan 1)" }
			"MCOSTANDARD_MIDMARKET" { $licenses += "Lync Online (Plan 1)" }
			"SHAREPOINTENTERPRISE_MIDMARKET" { $licenses += "Sharepoint Online (Plan 1)" }
			"SHAREPOINT_WAC" { $licenses += "Office Online" }
			"OFFICESUBSCRIPTION" { $licenses += "Office ProPlus" }
			"YAMMER_MIDSIZE" { $licenses += "Yammer" }
			"EXCHANGE_S_STANDARD" { $licenses += "Exchange Online (Plan 2)" }
			"MCOSTANDARD" { $licenses += "Lync Online (Plan 2)" }
			"SHAREPOINTENTERPRISE" { $licenses += "Sharepoint Online (Plan 2)" }
			"RMS_S_ENTERPRISE" { $licenses += "Azure Active Directory Rights Management" }
			"YAMMER_ENTERPRISE" { $licenses += "Yammer" }
			"MCVOICECONF" { $licenses += "Lync Online (Plan 3)" }
			"EXCHANGE_S_DESKLESS" { $licenses += "Exchange Online Kiosk" }
			"SHAREPOINTDESKLESS" { $licenses += "SharePoint Online Kiosk" }
			"STANDARDPACK_STUDENT" { $licenses += "Microsoft Office 365 (Plan A1) for Students" }
			"STANDARDPACK_FACULTY" { $licenses += "Microsoft Office 365 (Plan A1) for Faculty" }
			"STANDARDWOFFPACK_FACULTY" { $licenses += "Office 365 Education E1 for Faculty" }
			"STANDARDWOFFPACK_STUDENT" { $licenses += "Microsoft Office 365 (Plan A2) for Students" }
			"STANDARDWOFFPACK_IW_STUDENT" { $licenses += "Office 365 Education for Students" }
			"STANDARDWOFFPACK_IW_FACULTY" { $licenses += "Office 365 Education for Faculy" }
			"EOP_ENTERPRISE FACULTY" { $licenses += "Exchange Online Protection for Faculty" }
			"EXCHANGESTANDARD_STUDENT" { $licenses += "Exchange Online (Plan 1) for Students" }
			"OFFICESUBSCRIPTION_STUDENT" { $licenses += "Office ProPlus Student Benefit" }
			"OFFICESUBSCRIPTION_FACULTY" { $licenses += "Office ProPlus Faculty Benefit" }
			"PROJECTONLINE_PLAN1_FACULTY" { $licenses += "Project Online for Faculty" }
			"SHAREPOINTWAC_EDU" { $licenses += "Office Online for Education" }
			"SHAREPOINTENTERPRISE_EDU" { $licenses += "SharePoint Plan 2 for EDU" }
			"SHAREPOINT_PROJECT_EDU" { $licenses += "Project Online for Education" }
			"PROJECTONLINE_PLAN1_STUDENT" { $licenses += "Project Online for Students" }
			"STANDARDPACK_GOV" { $licenses += "Microsoft Office 365 (Plan G1) for Government" }
			"STANDARDWOFFPACK_GOV" { $licenses += "Microsoft Office 365 (Plan G2) for Government" }
			"ENTERPRISEPACK_GOV" { $licenses += "Microsoft Office 365 (Plan G3) for Government" }
			"ENTERPRISEWITHSCAL_GOV" { $licenses += "Microsoft Office 365 (Plan G4) for Government" }
			"DESKLESSPACK_GOV" { $licenses += "Microsoft Office 365 (Plan K1) for Government" }
			"EXCHANGESTANDARD_GOV" { $licenses += "Microsoft Office 365 Exchange Online (Plan 1) only for Goverment" }
			"EXCHANGEENTERPRISE_GOV" { $licenses += "Microsoft Office 365 Exchange Online (Plan 2) only for Goverment" }
			"SHAREPOINTDESKLESS_GOV" { $licenses += "SharePoint Online Kiosk" }
			"EXCHANGE_S_DESKLESS_GOV" { $licenses += "Exchange Kiosk" }
			"RMS_S_ENTERPRISE_GOV" { $licenses += "Windows Azure Active Directory Rights Management" }
			"OFFICESUBSCRIPTION_GOV" { $licenses += "Office ProPlus" }
			"MCOSTANDARD_GOV" { $licenses += "Lync Plan 2G" }
			"SHAREPOINTWAC_GOV" { $licenses += "Office Online for Government" }
			"SHAREPOINTENTERPRISE_GOV" { $licenses += "SharePoint Plan 2G" }
			"EXCHANGE_S_ENTERPRISE_GOV" { $licenses += "Exchange Plan 2G" }
			"EXCHANGE_S_ARCHIVE_ADDON_GOV" { $licenses += "Exchange Online Archiving" }
			"LITEPACK" { $licenses += "Office 365 (Plan P1)" }
			"STANDARDPACK" { $licenses += "Microsoft Office 365 (Plan E1)" }
			"STANDARDWOFFPACK" { $licenses += "Microsoft Office 365 (Plan E2)" }
			"DESKLESSPACK" { $licenses += "Office 365 (Plan K1)" }
			"EXCHANGEARCHIVE" { $licenses += "Exchange Online Archiving" }
			"EXCHANGETELCO" { $licenses += "Exchange Online POP" }
			"SHAREPOINTSTORAGE" { $licenses += "SharePoint Online Storage" }
			"SHAREPOINTPARTNER" { $licenses += "SharePoint Online Partner Access" }
			"PROJECTONLINE_PLAN_1" { $licenses += "Project Online (Plan 1)" }
			"PROJECTONLINE_PLAN_2" { $licenses += "Project Online (Plan 2)" }
			"PROJECT_CLIENT_SUBSCRIPTION" { $licenses += "Project Pro for Office 365" }
			"VISIO_CLIENT_SUBSCRIPTION" { $licenses += "Visio Pro for Office 365" }
			"INTUNE_A" { $licenses += "Intune for Office 365" }
			"CRMSTANDARD" { $licenses += "CRM Online" }
			"CRMTESTINSTANCE" { $licenses += "CRM Test Instance" }
			"ONEDRIVESTANDARD" { $licenses += "OneDrive" }
			"WACONEDRIVESTANDARD" { $licenses += "OneDrive Pack" }
			"SQL_IS_SSIM" { $licenses += "Power BI Information Services" }
			"BI_AZURE_P1" { $licenses += "Power BI Reporting and Analytics" }
			"EOP_ENTERPRISE" { $licenses += "Exchange Online Protection" }
			"PROJECT_ESSENTIALS" { $licenses += "Project Lite" }
			"CRMIUR" { $licenses += "CRM for Partners" }
			"NBPROFESSIONALFORCRM" { $licenses += "Microsoft Social Listening Professional" }
			"AAD_PREMIUM" { $licenses += "Azure Active Directory Premium" }
			"MFA_PREMIUM" { $licenses += "Azure Multi-Factor Authentication" }
			"CRMSTORAGE" { $licenses += "Microsoft Dynamics CRM Online Additional Storage" }
			"MDM_SALES_COLLABORATION" { $licenses += "Microsoft Dynamics Marketing Sales Collaboration" }
			"POWER_BI_STANDARD" { $licenses += "Power BI" }
			"O365_BUSINESS" { $licenses += "Office 365 Business" }
			"O365_BUSINESS_ESSENTIALS" { $licenses += "Office 365 Business Essentials" }
			"O365_BUSINESS_PREMIUM" { $licenses += "Office 365 Business Premium" }
			"SMB_BUSINESS" { $licenses += "Office 365 Business" }
			"SMB_BUSINESS_ESSENTIALS" { $licenses += "Office 365 Business Essentials" }
			"SMB_BUSINESS_PREMIUM" { $licenses += "Office 365 Business Premium" }
			"EXCHANGESTANDARD" { $licenses += "Exchange Online (Plan 1)" }
			"ENTERPRISEPACK" { $licenses += "Office 365 Enterprise E3" }
			"EXCHANGEENTERPRISE" { $licenses += "Exchange Online (Plan 2)" }
			"VISIOCLIENT" { $licenses += "Visio Pro for Office 365" }
			"ENTERPRISEPREMIUM" { $licenses += "Office 365 Enterprise E5" }
			default { $licenses += $sku }
		}
	}
	$licenses = $licenses -join "; "

	$exportData += "`"$userPrincipalName`",`"$displayName`",`"$office`",`"$lastLogonTime`",`"$accountEnabled`",`"$passwordExpiration`",`"$licenses`"`n"
}

# Convert data to Excel
$exportData = $exportData | ConvertFrom-CSV
$excel = New-Object -ComObject Excel.Application 
$excel.visible = $false
$excel.DisplayAlerts = $false
$workbooks = $excel.Workbooks.Add()
$worksheets = $workbooks.worksheets
$worksheet = $worksheets.Item(1)
$worksheet.Name = "Office 365"

# Add headers
$worksheet.Cells.Item(1,1) = "Username"
$worksheet.Cells.Item(1,2) = "Displayname"
$worksheet.Cells.Item(1,3) = "Office"
$worksheet.Cells.Item(1,4) = "Last login"
$worksheet.Cells.Item(1,5) = "Sign-in status"
$worksheet.Cells.Item(1,6) = "Password Expires"
$worksheet.Cells.Item(1,7) = "Licenses"

# Add data
$i = 2
foreach($row in $exportData)
{
	$worksheet.Cells.Item($i,1) = $row.UserPrincipalName
	$worksheet.Cells.Item($i,2) = $row.DisplayName
	$worksheet.Cells.Item($i,3) = $row.Office
	$worksheet.Cells.Item($i,4) = $row.LastLogonTime
	$worksheet.Cells.Item($i,5) = $row.AccountEnabled
	$worksheet.Cells.Item($i,6) = $row.PasswordNeverExpires
	$worksheet.Cells.Item($i,7) = $row.Licenses
	$i++
}

# Autofit all data
#$worksheet.columns.item("d").NumberFormat = "dd-mm-yyyy hh:mm"
$worksheet.UsedRange.EntireColumn.AutoFit()

# Format data as table
$excelTable = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $worksheet.UsedRange, 0, 1)
$excelTable.Name = "Table1"
$excelTable.TableStyle = "TableStyleMedium2"

# Prompt to save file
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveFileDialog.filter = "Excel Workbook|*.xlsx"
$saveFileDialog.ShowDialog() | Out-Null
if ($saveFileDialog.FileName -ne "") {
	$workbooks.saveas($saveFileDialog.FileName)
}

$excel.Quit()
Remove-Variable -Name excel
[gc]::collect() 
[gc]::WaitForPendingFinalizers()

# Cleanup the session to Exchange Online
Get-PSSession | Remove-PSSession
