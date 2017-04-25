# Create session to Office 365 and Exchange Online
Write-Host "Connecting to Office 365 and Exchange Online..."
$userCredential = Get-Credential -Credential $null
if (!$userCredential) { exit }
Connect-MsolService -Credential $userCredential
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $userCredential -Authentication Basic -AllowRedirection -ErrorAction Stop
Import-PSSession $session -ErrorAction Stop -DisableNameChecking -AllowClobber | Out-Null

# Export data from Office 365
$users = @()
$mailboxUsers = Get-MsolUser -All | Where { $_.IsLicensed -eq $true }

Write-Host "Exporting Office 365 user properties..." -ForegroundColor Yellow
foreach ($user in $mailboxUsers)
{
    $upn = $user.UserPrincipalName
    $username = $user.Name
    $mol = Get-MsolUser -UserPrincipalName $upn | Select-Object City, Country, Department, DisplayName, Fax, FirstName, LastName, MobilePhone, Office, PasswordNeverExpires, PhoneNumber, PostalCode,SignInName, State, StreetAddress, Title
    $emailAddress = Get-Mailbox -ResultSize Unlimited -Identity $upn -ErrorAction SilentlyContinue | Select-Object Name, DisplayName, PrimarySmtpAddress, EmailAddresses

    $properties = @{
      Name = $emailAddress.name
      City = $mol.City
      Country = $mol.Country
      Department = $mol.Department
      Displayname = $mol.DisplayName
      Emailaddress = $EmailAddress.PrimarySmtpAddress
      Fax = $mol.Fax
      FirstName = $mol.FirstName
      LastName = $mol.LastName
      MobilePhone = $mol.MobilePhone
      Office = $mol.Office
      PasswordNeverExpires = $mol.PasswordNeverExpires
      PhoneNumber = $mol.PhoneNumber
      PostalCode = $mol.PostalCode
      SignInName = $mol.SignInName
      State = $mol.State
      StreetAddress = $mol.StreetAddress
      Title = $mol.Title
      UserPrincipalName = $upn
      ProxyAddresses = $emailAddress.EmailAddresses
    }

    $users += New-Object PSObject -Property $properties
}

$adUsers = Get-ADUser -Filter *
foreach ($user in $users)
{
	# Get the username
    $signInName = $user.SignInName.Split("@")[0]
	
	# Check if the user exists
    $userExists = $adUsers | Where { $_.samAccountName -eq $signInName }
	
	# Update the attributes of the user with the attributes of Office 365
    if ($userExists)
    {
        Write-Host "Updating user $($signInName) with Office 365 values..."
        Set-ADUser -Identity $signInName `
            -SamAccountName $signInName `
            -GivenName $user.FirstName `
            -Surname $user.LastName `
            -City $user.City `
            -Department $user.Department `
            -DisplayName $user.DisplayName `
            -EmailAddress $user.EmailAddress `
            -Fax $user.Fax `
            -MobilePhone $user.MobilePhone `
            -Office $user.Office `
            -OfficePhone $user.PhoneNumber `
            -PostalCode $user.PostalCode `
            -State $user.State `
            -StreetAddress $user.StreetAddress `
            -Title $user.Title `
            -UserPrincipalName $user.UserPrincipalName

        foreach ($address in $user.ProxyAddresses)
        {
            if ($address -ne "")
            {
                Set-ADUser -Identity $signInName -Add @{ProxyAddresses=$Address}
            }
        }
    }
    else
    {
        Write-Host "$signInName does not exist in Active Directory, creating it now..." -ForegroundColor Yellow
        try
        {
            New-ADUser -Name $user.Displayname `
                -SamAccountName $signInName `
                -GivenName $user.FirstName `
                -Surname $user.LastName `
                -City $user.City `
                -Department $user.Department `
                -DisplayName $user.DisplayName `
                -EmailAddress $user.EmailAddress `
                -Fax $user.Fax `
                -MobilePhone $user.MobilePhone `
                -Office $user.Office `
                -OfficePhone $user.PhoneNumber `
                -PostalCode $user.PostalCode `
                -State $user.State `
                -StreetAddress $user.StreetAddress `
                -Title $user.Title `
                -UserPrincipalName $user.UserPrincipalName `
                -Path "OU=de-DE,OU=IT_PE_Users,OU=IT_PE,DC=permadental,DC=local" `
                -AccountPassword ("Welkom123!" | ConvertTo-SecureString -AsPlainText -Force) `
                -Enabled $true `
                -ErrorAction Stop

            foreach ($address in $user.ProxyAddresses)
            {
                if ($address -ne "")
                {
                    Set-ADUser -Identity $signInName -Add @{ProxyAddresses=$address}
                }
            }
        }
        catch
        {
            Write-Host "Error creating user $($user.DisplayName)" -ForegroundColor Red
        }
    }
}

Write-Host "Updating ImmutableId's in Office 365..." -ForegroundColor Yellow
foreach ($user in $users)
{
    # Get the username
    $signInName = $user.SignInName.Split("@")[0]
	
	# Check if the user exists
    $userExists = $adUsers | Where { $_.samAccountName -eq $signInName }

    if ($userExists)
    {
        $adUser = Get-ADUser $signInName -Properties *
        $immutableId = [system.convert]::ToBase64String(([GUID]($adUser.ObjectGUID)).tobytearray())

        Set-MsolUser -UserPrincipalName $user.UserPrincipalName -ImmutableId $immutableId
    }
}

Write-Host "Exporting Office 365 group properties..." -ForegroundColor Yellow
$groups = Get-MsolGroup -All

$adGroups = Get-ADGroup -Filter * -Properties *
foreach ($group in $groups)
{
    # Check if the group exists
    $adGroup = $null
    if ($group.GroupType -eq "DistributionList") { $adGroup = $adGroups | Where { $_.GroupCategory -eq "Distribution" -and $_.Name -eq $group.DisplayName} }
    if ($group.GroupType -eq "MailEnabledSecurity") { $adGroup = $adGroups | Where { $_.GroupCategory -eq "Security" -and $_.Name -eq $group.DisplayName} }
    
    if ($adGroup)
    {
        Write-Host "Updating group $($group.DisplayName) with Office 365 values..."

        Set-ADGroup -Identity $adGroup.SamAccountName -DisplayName $group.DisplayName

        Set-ADGroup -Identity $adGroup.SamAccountName -Replace @{mail=$group.EmailAddress}

        foreach ($address in $group.ProxyAddresses)
        {
            if ($address -ne "")
            {
                Set-ADGroup -Identity $adGroup.SamAccountName -Add @{proxyAddresses=$address}
            }
        }

        if ($group.GroupType -eq "DistributionList")
        {
            $groupMembers = Get-DistributionGroupMember -Identity $group.EmailAddress
            foreach ($member in $groupMembers)
            {
                $username = $member.WindowsLiveID.Split("@")[0]
                Add-ADGroupMember -Identity $adGroup.SamAccountName -Members $username
            }
        }
    }
}

# Cleanup session
Get-PSSession | Remove-PSSession
