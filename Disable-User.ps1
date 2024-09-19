$ADFieldsGet = @(
    'samAccountName'
    'mail'
    'name'
    'userPrincipalName'
    'telephoneNumber'
    'mobile'
    'title'
    'company'
    'description'
    'physicalDeliveryOfficeName'
    'department'
    'enabled'
    'manager'
    'streetAddress'
    'l'
    'state'
    'postalCode'
    'memberOf'
    'canonicalName'
    'directReports'
    )

$ADFieldsDisplay = @(
    @{n='Full Name';e={($_.Name)}}
    @{n='Username';e={($_.SamAccountName)}}
    @{n='Email';e={($_.UserPrincipalName)}}
    @{n='Telephone Number';e={($_.Telephonenumber)}}
    @{n='Mobile';e={($_.Mobile)}}
    @{n='Title';e={($_.Title)}}
    @{n='Company';e={($_.Company)}}
    @{n='Description';e={($_.Description)}}   
    @{n='Office';e={($_.physicalDeliveryOfficeName)}}
    @{n='Department';e={($_.Department)}}
    @{n='Enabled';e={($_.Enabled)}}
    @{n='Manager';e={($_.Manager.split(',')| Where-Object {$_.StartsWith("CN=")}) -replace "CN="}}
    @{n='Address';e={($_.StreetAddress)}}
    @{n='State';e={($_.st)}}
    @{n='Country';e={($_.co)}} 
    @{n='Zip Code';e={($_.PostalCode)}}
    @{n='Groups';e={($_.MemberOf.split(',')| Where-Object {$_.StartsWith("CN=")}) -join "`n" -replace "CN="}}
    @{n='OU';e={($_.CanonicalName)}}
    @{n='Direct Reports';e={($_.DirectReports.split(',') | Where-Object {$_.StartsWith("CN=")}) -join "`n" -replace "CN="}}
    )

$ADFieldsClear = @(
    'TelephoneNumber'
    'Mobile'
    'Title'
    'Company'
    'Description'
    'physicalDeliveryOfficeName'
    'Department'
    'Manager'
    'L'
    'St'
    'StreetAddress'
    'PostalCode'
    )

# Gets the User's and Manager's info from AD
$username = Read-Host -Prompt "Please enter the username of the terminated account"
$user = Get-ADUser -Identity $username -Properties $ADFieldsGet 
$Mgr = Get-ADuser $user.Manager -Properties SamAccountName,mail
$Message = "$($user.SamAccountName) is no longer with the company. Please contact $($Mgr.SamAccountName) at $($Mgr.mail)"
$ExchangeCreds = Get-Credential -Message "Enter your Email and password (for Exchange Management)"

# Displays information from AD about the user
function Show-User-Info {
    Write-Host "Active Directory Fields" -BackgroundColor DarkGreen -ForegroundColor Black
    $user = Get-ADUser -Identity $username -Properties $ADFieldsGet | Select $ADFieldsDisplay
    $user
}

# Confirms if the selected user is correct. 
function Confirm-User {
    $continue = Read-Host "User found! Would you like to continue? (y/n)"
    if ($continue -eq "y" -or $continue -eq "Y") {
        
    } elseif ($continue -eq "n" -or $continue -eq "N") {
        Write-Host "Exiting..."
        Break
    }
}

# Removes user from groups
function Remove-Groups {
    Write-Host "Removing user from groups..." -ForegroundColor Yellow
    Get-ADUser -Identity $user -Properties MemberOf | ForEach-Object {
      $_.MemberOf | Remove-ADGroupMember -Members $_.SamAccountName -Confirm:$false -PassThru:$true | Select @{n='Groups';e={($_.DistinguishedName.split(',')| Where-Object {$_.StartsWith("CN=")}) -join "`n" -replace "CN="}}
    }
    $user = Get-ADUser -Identity $user -Properties MemberOf | Select MemberOf
    if ([string]::IsNullOrWhitespace($user.MemberOf)) {
        Write-Host "Groups Removed!" -ForegroundColor Cyan
        Write-Host $ADFieldsDisplay.Groups
    } else {
        Write-Host "Groups not removed! Please remove them manually." -ForegroundColor Red
    }
}
# Disables Account in AD
function Disable-User {
    Write-Host "Diabling account in AD..." -ForegroundColor Yellow
    Disable-ADAccount -Identity $user
    $user = Get-ADUser -Identity $username -Properties Enabled | Select Enabled
    if (!$user.Enabled) {
        Write-Host "Account disabled!" -ForegroundColor Cyan
    } else {
        Write-Host "Account not disabled! Please disable manually." -ForegroundColor Red
    }
}

# Moves user to terminated OU
function Move-To-Temp-OU {
    Write-Host "Moving user to Terminated OU..." -ForegroundColor Yellow
    Move-ADObject -Identity $user -TargetPath "<Distinguished path to  Terminated OU>"
    $user = Get-ADUser -Identity $username -Properties CanonicalName | select @{n='OU';e={($_.CanonicalName  -split "/")[-2]}}
    if ($user.OU -eq "Terminated OU") {
        Write-Host "Moved to Term OU!" -ForegroundColor Cyan
    } else {
        Write-Host "Not moved to Terminated OU! Please move the account manually." -ForegroundColor Red
    }
} 

# Clears user's AD fields
function Clear-AD-Fields {
    Write-Host "Clearing AD fields..." -ForegroundColor Yellow
    Set-ADUser -Identity $username -Clear $ADFieldsClear
    $user = Get-ADUser -Identity $username -Properties $ADFieldsGet | select $ADFieldsClear
 
    if ([string]::IsNullOrWhitespace($user.$ADFieldsClear)) {
            Write-Host "AD fields cleared!" -ForegroundColor Cyan
        } else {        
            Write-Host "AD fields not cleared! Please clear fields manually." -ForegroundColor Red
        }

}
# Connects to the Exchange Online module, then Sets Autoreply, converts to shared mailbox, and enables forwarding to manager
function Update-Email {
    Write-Host "Connecting to Microsoft Exchange..." -ForegroundColor Yellow
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -Credentials -ShowBanner:$false
    Write-Host "Connected!" -ForegroundColor Cyan
    Write-Host "Adding Autoreply for $($user.mail)..." -ForegroundColor Yellow
    Set-MailboxAutoReplyConfiguration -Identity $user.mail -AutoReplyState Enabled -InternalMessage $Message -ExternalMessage $Message -ExternalAudience All
    Write-Host "Autoreply added!" -ForegroundColor Cyan
    Write-Host "Forwarding email to $($Mgr.mail)..." -ForegroundColor Yellow
    Set-Mailbox -Identity $user.mail -ForwardingAddress $Mgr.mail
    Write-Host "Forwarding enabled!" -ForegroundColor Cyan
    Write-Host "Converting user mailbox to a shared mailbox..." -ForegroundColor Yellow
    Set-Mailbox -Identity $user.mail -Type Shared
    Write-Host "Converted to shared mailbox!" -ForegroundColor Cyan
}

# Informs that the termination is complete and displays properties to verify
function Term-Finished {
    Write-Host "Termination complete! Please verify the information below..."-ForegroundColor Green
    Write-Host "Exchange information:" -BackgroundColor White -ForegroundColor Black
    Get-MailboxAutoReplyConfiguration -Identity $user.mail | select AutoReplyState, InternalMessage, ExternalMessage, ExternalAudience | Format-List 
    Show-User-Info
    Disconnect-ExchangeOnline -Confirm:$false
}

Show-User-Info
Confirm-User
Remove-Groups
Disable-User
Clear-AD-Fields
Update-Email
Move-To-Temp-OU
Term-Finished

