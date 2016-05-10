# TODO: Fix and complete this script
# Create session to Office 365 and Exchange Online
Write-Host "Connecting to Office 365 and Exchange Online..."
$UserCredential = Get-Credential -Credential $null
if (!$UserCredential) { exit }
Connect-MsolService -Credential $UserCredential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection -ErrorAction Stop
Import-PSSession $Session -ErrorAction Stop -DisableNameChecking -AllowClobber | Out-Null

# Generate random password
Function Random-Password ($length = 8)
{
    $punc = 46..46
    $digits = 48..57
    $letters = 65..90 + 97..122

    $password = get-random -count $length `
        -input ($punc + $digits + $letters) |
            % -begin { $aa = $null } `
            -process {$aa += [char]$_} `
            -end {$aa}

    return $password
}

# Export data from Office 365
$Results = @()
$MailboxUsers = Get-MsolUser -All | Where { $_.IsLicensed -eq $true }

foreach($User in $MailboxUsers)
{
    Write-Host "Exporting data for user $($User.UserPrincipalName)"

    $UPN = $User.UserPrincipalName
    $Username = $User.Name
    $MOL = Get-MsolUser -UserPrincipalName $UPN | Select-Object City, Country, Department, DisplayName, Fax, FirstName, LastName, MobilePhone, Office, PasswordNeverExpires, PhoneNumber, PostalCode,SignInName, State, StreetAddress, Title
    $EmailAddress = Get-Mailbox -ResultSize Unlimited -Identity $UPN -ErrorAction SilentlyContinue | Select-Object Name, DisplayName, PrimarySmtpAddress, EmailAddresses

    $ProxyAddresses = ""
    foreach ($ProxyAddress in $EmailAddress.EmailAddresses)
    {
        $ProxyAddresses += $ProxyAddress + ";"
    }
    $ProxyAddresses.Substring(0,$ProxyAddresses.Length-1) | Out-Null

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
      Password = Random-Password
      ProxyAddresses = $ProxyAddresses
    }

    $Results += New-Object PSObject -Property $Properties
}

$Results | Select-Object Name, City, Country, Department, DisplayName, Emailaddress, Fax, FirstName, LastName, MobilePhone, Office, PasswordNeverExpires, PhoneNumber, PostalCode,SignInName, State, StreetAddress, Title, UserPrincipalName, Password, ProxyAddresses | Export-Csv -Path ".\Office365Export.csv" -Encoding UTF8 -NoTypeInformation

Get-PSSession | Remove-PSSession
$Users = Import-Csv ".\Office365Export.csv" -Encoding UTF8
$AdUsers = Get-ADUser -Filter *
foreach ($User in $Users)
{
    $SignInName = $User.SignInName.Split("@")[0]
    Write-Host "$SignInName"

    $ProxyAddresses = $User.ProxyAddresses -Split ";"

    $CheckIfUserExists = ($AdUsers | Where { $_.samAccountName -eq $SignInName }) -gt 0

    if ($CheckIfUserExists)
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
            -PasswordNeverExpires $True `
            -OfficePhone $User.PhoneNumber `
            -PostalCode $User.PostalCode `
            -State $User.State `
            -StreetAddress $User.StreetAddress `
            -Title $User.Title `
            -UserPrincipalName $User.UserPrincipalName
    }
    else
    {
        Write-Host "$($SignInName) does not exist in Active Directory, skipping until finished coding :)" -ForegroundColor Red
    }
}