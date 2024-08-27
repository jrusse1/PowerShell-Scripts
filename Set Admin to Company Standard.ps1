$KeyFile = "C:\PATH\TO\AES.key"
$PasswordFile = "C:\PATH\TO\ENCRYPTED\Password.txt"
$Key = Get-Content $KeyFile
$EncryptedPassword = Get-Content $PasswordFile
$SecurePassword = ConvertTo-SecureString -String $EncryptedPassword -Key $Key

#Creates a list for all users on the computer
$list = @(Get-WmiObject -Class Win32_UserAccount -Filter 'LocalAccount=true' | Select name)

#Loops through the list to determine if <NEW ADMIN NAME> is already present, and will add it to the Administrators group
If ($list.name -contains "<NEW ADMIN NAME>") {
    Add-LocalGroupMember -Group "Administrators" -Member "<NEW ADMIN NAME>"
    Continue

} else { 
    foreach ($element in $list) {

#Checks if account is using outdated name, changes it to "<NEW ADMIN NAME" and adds it to the Administrator group
        if ($element.name -eq "<OLD ADMIN NAME 1>" -or $element.name -eq "<OLD ADMIN NAME 2>" -or $element.name -eq '<OLD ADMIN NAME 3>'){
            Rename-LocalUser -Name $element.name -NewName "<NEW ADMIN NAME>" -ErrorAction SilentlyContinue
            Set-localuser -Name "<NEW ADMIN NAME>" -Password $SecurePassword -ErrorAction SilentlyContinue
    
        } else {
            continue

        }
    }
}

#If no company standard Administrator account is found above, creates a new Local Administrator account called "<NEW ADMIN NAME>"
if ($list.name -notcontains "<NEW ADMIN NAME>"){
    New-LocalUser -Name "<NEW ADMIN NAME>" -Password $SecurePassword 
    Add-LocalGroupMember -Group "Administrators" -Member "<NEW ADMIN NAME>" 
}

        
