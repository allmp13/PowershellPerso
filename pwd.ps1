<#$Credentials = Get-Credential
$Credentials.Password | ConvertFrom-SecureString | Set-Content C:\password.txt#>




$username = "-Admin-jmdavid"
$password = Get-Content C:\password.txt | convertto-securestring
$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $password

$NewPasswd = ConvertTo-SecureString -AsPlainText "Azerty1" -force

Get-ADUser -identity nabbes -properties EmployeeNumber
Set-ADAccountPassword -Credential $cred -Identity usert -Reset -NewPassword $NewPasswd