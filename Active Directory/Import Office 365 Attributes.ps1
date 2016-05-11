# Create session to Office 365 and Exchange Online
Write-Host "Connecting to Office 365 and Exchange Online..."
$UserCredential = Get-Credential -Credential $null
if (!$UserCredential) { exit }
Connect-MsolService -Credential $UserCredential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection -ErrorAction Stop
Import-PSSession $Session -ErrorAction Stop -DisableNameChecking -AllowClobber | Out-Null

# Export data from Office 365
$Results = @()
$MailboxUsers = Get-MsolUser -All | Where { $_.IsLicensed -eq $true }

Foreach ($User in $MailboxUsers)
{
    Write-Host "Exporting data for user $($User.UserPrincipalName)"

    $UPN = $User.UserPrincipalName
    $Username = $User.Name
    $MOL = Get-MsolUser -UserPrincipalName $UPN | Select-Object City, Country, Department, DisplayName, Fax, FirstName, LastName, MobilePhone, Office, PasswordNeverExpires, PhoneNumber, PostalCode,SignInName, State, StreetAddress, Title
    $EmailAddress = Get-Mailbox -ResultSize Unlimited -Identity $UPN -ErrorAction SilentlyContinue | Select-Object Name, DisplayName, PrimarySmtpAddress, EmailAddresses

    $Properties = @{
      Name = $EmailAddress.name
      City = $MOL.City
      Country = $MOL.Country
      Department = $MOL.Department
      Displayname = $MOL.DisplayName
      Emailaddress = $EmailAddress.PrimarySmtpAddress
      Fax = $MOL.Fax
      FirstName = $MOL.FirstName
      LastName = $MOL.LastName
      MobilePhone = $MOL.MobilePhone
      Office = $MOL.Office
      PasswordNeverExpires = $MOL.PasswordNeverExpires
      PhoneNumber = $MOL.PhoneNumber
      PostalCode = $MOL.PostalCode
      SignInName = $MOL.SignInName
      State = $MOL.State
      StreetAddress = $MOL.StreetAddress
      Title = $MOL.Title
      UserPrincipalName = $UPN
      ProxyAddresses = $EmailAddress.EmailAddresses
    }

    $Results += New-Object PSObject -Property $Properties
}

# Cleanup session
Get-PSSession | Remove-PSSession

$AdUsers = Get-ADUser -Filter *
Foreach ($User in $Results)
{
	# Get the username
    $SignInName = $User.SignInName.Split("@")[0]
	
	# Check if the user exists
    $UserExists = ($AdUsers | Where { $_.samAccountName -eq $SignInName }) -gt 0
	
	# Update the attributes of the user with the attributes of Office 365
    if ($UserExists)
    {
        Write-Host "Updating user $($SignInName) with Office 365 values..."
        Set-ADUser -Identity $SignInName `
            -Name $User.Displayname `
            -SamAccountName $SignInName `
            -GivenName $User.FirstName `
            -Surname $User.LastName `
            -City $User.City `
            -Department $User.Department `
            -DisplayName $User.DisplayName `
            -EmailAddress $User.EmailAddress `
            -Fax $User.Fax `
            -MobilePhone $User.MobilePhone `
            -Office $User.Office `
            -OfficePhone $User.PhoneNumber `
            -PostalCode $User.PostalCode `
            -State $User.State `
            -StreetAddress $User.StreetAddress `
            -Title $User.Title `
            -UserPrincipalName $User.UserPrincipalName
    }
    else
    {
        Write-Host "$SignInName does not exist in Active Directory" -ForegroundColor Red
    }
}