Clear-Host
$password = Cat c:\scripts\MonFichier.pwd | ConvertTo-SecureString
$user = "ddsis-13\jmdavid"
$credential = New-Object System.Management.Automation.PSCredential($user,$password)
$session = New-PSSession -ConfigurationName Microsoft.Exchange –ConnectionUri http://srvcas1/powershell -Authentication Kerberos -Credential $credential
$output=Import-PSSession $session -AllowClobber -DisableNameChecking 

(Get-Recipient -Identity "jmdavid@sdis13.fr").RecipientType

Remove-PSSession $session
