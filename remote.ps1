Clear-Host
$password = Cat c:\scripts\MonFichier.pwd | ConvertTo-SecureString
$user = "ddsis-13\jmdavid"
$credential = New-Object System.Management.Automation.PSCredential($user,$password)
$session = New-PSSession -ConfigurationName Microsoft.Exchange –ConnectionUri http://srvcas1/powershell -Credential $credential
Import-PSSession $session
Get-DistributionGroupMember -Identity aixinf@sdis13.fr
Remove-PSSession $session
