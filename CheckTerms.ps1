<# 
This program reads a spreadsheet of terminated users and checks Active Directory to see if they are terminated.
It will export each user to a new spreadsheet showing all of the data that should have been cleared, as it's presented in Active Directory.

The spreadsheet will need to be passed as an argument when the script is run.

Example: 
.\CheckTerms.ps1 "Terminations.xlsx"
#>

#List of the fields that should be cleared
$ADFields = @(
    'samAccountName'
    'Enabled'
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

#Reads the file in the argument
$excelFileIn = Import-Excel $args[0]
$excelFileOut = "C:\Users\jrussell\Termination Lists\" + (Get-Date -Format "MM-dd-yyyy-HH-mm-ss") + ".xlsx"
$dataOut = @()

#Removes email domain and searches Active Directory for the user
foreach ($user in $excelFileIn) {
    try {
        if ($null -eq $user.Email) {
            Continue
        }

        else {
            $email = $user.Email
            $username = $email.Split('@')[0]
            $userData = Get-ADUser -Identity $username -properties $ADFields | Select-Object $ADFields
            $dataOut += $userData
        }
 
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        
        "Cannot find user: $username!"
       }
    }

$XL = $dataOut | Export-Excel -Path $excelFileOut -AutoSize -FreezeTopRow -WorksheetName "Sheet1" -PassThru
Add-ConditionalFormatting -WorkSheet $XL.Workbook.Worksheets[1] -Address "a2:b1048576" -BackgroundColor 'yellow' -RuleType ContainsText -ConditionValue "=TRUE"
Close-ExcelPackage -ExcelPackage $XL -Show

