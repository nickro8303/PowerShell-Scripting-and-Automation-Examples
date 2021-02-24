

clear

try {

## Create organizational unit finance and import user info from financepersonnel.csv

New-ADOrganizationalUnit -name finance -ProtectedFromAccidentalDeletion $false

$NewAD = Import-Csv $PSScriptRoot\financePersonnel.csv
$Path = "OU=finance,DC=ucertify,DC=com"

ForEach ($ADuser in $NewAD)
{

$First = $ADuser.First_Name
$Last = $ADuser.Last_Name
$Name = "$First, $Last"
$SamName = $ADuser.samAccount
$City = $ADuser.City
$PostCode = $ADuser.PostalCode
$Office = $ADuser.OfficePhone
$Mobile = $ADuser.MobilePhone

New-ADUser -Name $Name -GivenName $First -Surname $Last -SamAccountName $SamName -City $City -PostalCode $PostCode -OfficePhone $Office -MobilePhone $Mobile -Path $Path
}

Import-Module SQLPS -DisableNameChecking -Force

## Create SQL Database and import user data from NewClientData.csv

$srv = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList .\ucertify3
$db = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Database -ArgumentList $srv, ClientDB
$db.Create()

Invoke-Sqlcmd -serverinstance .\ucertify3 -database ClientDB -inputfile $PSScriptRoot\Client_A_Contacts.sql

$table = 'dbo.Client_A_Contacts'
$db = 'ClientDB'

Import-csv $PSScriptRoot\NewClientData.csv | ForEach-Object {Invoke-Sqlcmd -database ClientDB -serverinstance .\ucertify3 `
-query "insert into $table (first_name,last_name, city, county, zip, officePhone, mobilePhone) VALUES `
('$($_.first_name)','$($_.last_name)','$($_.city)','$($_.county)','$($_.zip)','$($_.officePhone)','$($_.mobilePhone)')"
}

    
}

catch [System.OutOfMemoryException] {
    Write-Host "A system out of memory exception has occured."
    }
