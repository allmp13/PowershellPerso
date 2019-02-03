Clear-Host
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
Write-Host "Current script directory is $ScriptDir"
Read-Host -prompt "Entrer un mot de passe" -AsSecureString | 
ConvertFrom-SecureString | out-file c:\scripts\MonFichier.pwd
$password = Cat c:\scripts\MonFichier.pwd | ConvertTo-SecureString
$user = "ddsis-13\jmdavid"

$credential = New-Object System.Management.Automation.PSCredential($user,$password)

