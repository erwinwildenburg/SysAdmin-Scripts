# Import required modules
Import-Module .\Shared\Get-GraphHeader.ps1
Import-Module .\Shared\Get-PartnerContracts.ps1
Import-Module .\Shared\Get-Users.ps1

# Get our own clientId
Get-GraphHeader | Out-Null
$clientId = ([Microsoft.IdentityModel.Clients.ActiveDirectory.TokenCache]::DefaultShared).ReadItems().ClientId

# Get all our partner tenants
$partners = Get-PartnerContracts -clientId $clientId

$report = @()
foreach ($partner in $partners)
{
    $users = Get-Users -clientId $clientId -tenantId $partner.customerId

    # Calculate the used licenses
    $exchangeUsed = 0
    $sharepointUsed = 0
    foreach ($user in $users)
    {
        $sharePointDone = $false
        $exchangeDone = $false
        foreach ($plan in ($user.assignedPlans | Where-Object { $_.capabilityStatus -eq "Enabled" }))
        {
            if ($plan.capabilityStatus -eq "Enabled")
            {
                if ($sharePointDone -ne $true -and $plan.service -eq "SharePoint")
                {
                    $sharepointUsed += 1
                    $sharePointDone = $true
                }
                elseif ($exchangeDone -ne $true -and $plan.service -eq "exchange")
                {
                    $exchangeUsed += 1
                    $exchangeDone = $true
                }
            }
        }
    }
    $report += @{
        Customer = $partner.displayName
        ExchangeUsage = $exchangeUsed
        SharePointUsage = $sharePointUsed
    }
}

$exportData = "Customer,ExchangeUsage,SharePointUsage`n"
foreach ($object in $reports)
{
    $exportData += "`"$($object.Customer)`",`"$($object.ExchangeUsage)`",`"$($object.SharePointUsage)`"`n"
}

# Convert data to Excel
$excel = New-Object -ComObject Excel.Application 
$excel.visible = $false
$excel.DisplayAlerts = $false
$workbooks = $excel.Workbooks.Add()
$worksheets = $workbooks.worksheets
$worksheet = $worksheets.Item(1)
$worksheet.Name = "Licenses"

# Add headers
$worksheet.Cells.Item(1,1) = "Customer"
$worksheet.Cells.Item(1,2) = "Exchange Usage"
$worksheet.Cells.Item(1,3) = "SharePoint Usage"

# Add data
$i = 2
foreach($object in $report)
{
	$worksheet.Cells.Item($i,1) = $object.Customer
	$worksheet.Cells.Item($i,2) = $object.ExchangeUsage
	$worksheet.Cells.Item($i,3) = $object.SharePointUsage
	$i++
}

# Autofit all data
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
