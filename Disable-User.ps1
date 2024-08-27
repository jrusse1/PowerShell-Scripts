$username = Read-Host -Prompt "Please enter the username of the terminated account"
$server = "<DOMAIN CONTROLLER>"
$exchangeemail = "<EMAIL WITH ADMIN ACCESS FOR EXCHANGE>"
$user = Get-ADUser -Identity $username -server $server -Properties *
$Mgr = Get-ADuser $user.Manager -Properties Displayname,EmailAddress
$MgrEmail = Get-ADUSer -Identity $Mgr.SamAccountName -Properties EmailAddress
$Message = "$($user.DisplayName) is no longer with the company. Please contact $($Mgr.DisplayName) at $($MgrEmail.EmailAddress)"

Write-Host "Active Directory Fields" -ForegroundColor Green
Get-ADUser -Identity $username -Properties * | select SamAccountName, Name, UserPrincipalName, title, Company, Department, Enabled, Manager, StreetAddress, PostalCode, MemberOf, DistinguishedName, directReports | Format-List

#Removes user from groups
Write-Host "Removing user from groups..." -ForegroundColor Yellow
Get-ADUser -Identity $user -Properties MemberOf | ForEach-Object {
  $_.MemberOf | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false
}

#Disables Account in AD
Write-Host "Diabling account in AD..." -ForegroundColor Yellow
Disable-ADAccount -Identity $user

#Moves user to term OU
Write-Host "Moving user to Term User Shared OU..." -ForegroundColor Yellow
Move-ADObject -Identity $user -TargetPath "OU=Term User Shared MB,OU=Forward Mailbox,OU=Email Accounts,OU=Service Accounts,DC=psmicorp,DC=com"

#Clears user's AD fields
Write-Host "Clearing AD fields..." -ForegroundColor Yellow
Set-ADUser -Identity $($user.SamAccountname) -Clear @('StreetAddress','l','st','title','physicalDeliveryOfficeName','o','telephoneNumber','postalcode','mobile','department','company','manager','description')

#Imports and connects Exchange Online Module
#Will need to be signed into Exchange SSO
Write-Host "Connecting to Microsoft Exchange..." -ForegroundColor Yellow
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName $exchangeemail -ShowBanner:$false

#Sets Autoreply, converts to shared mailbox, and enables forwarding to manager
Write-Host "Adding Autoreply for $($user.EmailAddress)..." -ForegroundColor Yellow
Set-MailboxAutoReplyConfiguration -Identity $user.EmailAddress -AutoReplyState Enabled -InternalMessage $Message -ExternalMessage $Message -ExternalAudience All
Write-Host "Forwarding email to $($MgrEmail.EmailAddress)..." -ForegroundColor Yellow
Set-Mailbox -Identity $user.EmailAddress -ForwardingAddress $MgrEmail.EmailAddress
Write-Host "Converting user mailbox to a shared mailbox..." -ForegroundColor Yellow
Set-Mailbox -Identity $user.EmailAddress -Type Shared

#Informs that the termination is complete and displays properties to verify
Write-Host "Termination complete! Please verify the information below..." -ForegroundColor Cyan
Write-Host "Exchange information:" -ForegroundColor Green
Get-MailboxAutoReplyConfiguration -Identity $user.EmailAddress | select AutoReplyState, InternalMessage, ExternalMessage, ExternalAudience | Format-List 
Write-Host "Active Directory Fields" -ForegroundColor Green
Get-ADUser -Identity $username -server $server -Properties * | select SamAccountName, Name, UserPrincipalName, title, Company, Department, Enabled, Manager, StreetAddress, PostalCode, MemberOf, DistinguishedName, directReports | Format-List

Disconnect-ExchangeOnline -Confirm:$false
