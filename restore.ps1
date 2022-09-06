# Tassy Estinphon #001392883
# Scripting and Automation Task 2

#Variable
$PSScriptRoot = "C:\Users\LabAdmin\Desktop\Requirements2"
    Try{ 
#Importing Modules
Import-Module ActiveDirectory
if (Get-Module -Name sqlps) {Remove-Module sqlps}
Import-Module -Name SqlServer
#Creating AD Organizational Unit
Write-Host -ForegroundColor Cyan "Starting Active Directory Tasks"
Write-Host -ForegroundColor Cyan "Creating Organizational Unit"
$OUCanName = "Finance"
$OUDName = "Fin"
New-ADOrganizationalUnit -Name $OUCanName -DisplayName $OUDName  -ProtectedFromAccidentalDeletion $false
#Importing Users
Write-Host -ForegroundColor Cyan "Importing Users"
$NewADUsers = Import-Csv -Path $PSScriptRoot\financepersonnel.csv
$path = "OU=Finance,DC=consultingfirm,DC=Com"
foreach ($ADUser in $NewADUsers)
{
    $First = $ADUser.First_Name
    $Last = $ADUser.Last_Name
    $Name = $First + " " + $Last
    $SamAcct = $ADUser.samAccount
    $Postal = $ADUser.PostalCode
    $Office = $ADUser.OfficePhone
    $Mobile = $ADUser.MobilePhone

    New-ADUser -GivenName $First -Surname $Last -Name $Name -SamAccountName $SamAcct -DisplayName $Name -PostalCode $Postal -MobilePhone $Mobile -OfficePhone $Office -Path $path
   }  
#Creating SQL Database
   $sqlServerInstanceName = ".\SQLEXPRESS"   
   $sqlServerObject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $sqlServerInstanceName
   $databaseName = 'ClientDB'
   $databaseObject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Database -ArgumentList $sqlServerObject, $databaseName 
   $databaseObject.Create()  
   
#Create Table
$tablename = "Client_A_Contacts"
Invoke-Sqlcmd -ServerInstance $sqlServerInstanceName -Database $databaseName -InputFile $PSScriptRoot\Client_A_Contacts.sql
#CSV File Info
$Insert = "INSERT INTO [$($tablename)] (first_name, last_name, city, county, zip, officePhone, mobilePhone) "
$NewClientData = Import-Csv $PSScriptRoot\NewClientData.csv

foreach ($NewClient in $NewClientData) 
{
    $Values = "Values ( `
                        '$($NewClient.first_name)', `
                        '$($NewClient.last_name)', `
                        '$($NewClient.city)', `
                        '$($NewClient.county)', `
                        '$($NewClient.zip)', `
                        '$($NewClient.officePhone)', `
                        '$($NewClient.mobilePhone)')"
    $query = $Insert + $Values
    Invoke-sqlcmd -Database $databaseName -ServerInstance $sqlServerInstanceName -Query $query  
}
#Catch Exem
    }catch [System.OutOfMemoryException]
{Write-Host = "A System Out of Memory Exception has occured"}