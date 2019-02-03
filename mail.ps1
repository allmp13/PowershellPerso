
<#
	SCRIPT : mail.ps1
#>
Clear-Host
$password = Cat c:\scripts\MonFichier.pwd | ConvertTo-SecureString
$user = "ddsis-13\jmdavid"
$credential = New-Object System.Management.Automation.PSCredential($user,$password)
$session = New-PSSession -ConfigurationName Microsoft.Exchange –ConnectionUri http://srvcas1/powershell -Credential $credential
Import-PSSession $session -AllowClobber

$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
Write-Host "Current script directory is $ScriptDir"

$Db = "c:\scripts\jmd.accdb"
$Tables = "matable"
$OpenStatic = 3
$LockOptimistic = 3
$connection = New-Object -ComObject ADODB.Connection
$connection.Open("Provider = Microsoft.ACE.OLEDB.12.0;Data Source=$Db" )
$RecordSet = new-object -ComObject ADODB.Recordset
ForEach($Table in $Tables)
     {
      $Query = "Select * from $Table"
      $RecordSet.Open($Query, $Connection, $OpenStatic, $LockOptimistic)
	  $fin = $RecordSet.Fields.Count
	  Write-Host $fin " Enregistrement(s)"
      $RecordSet.movefirst()
	  do 
	  	{
			for($i=0; $i -lt $fin; $i++){
			 Write-Host $Recordset.Fields.Item($i).Value
			
			}
			$s=Get-DistributionGroupMember -Identity $Recordset.Fields.Item(1).Value | select displayname, Name, PrimarySmtpAddress
			
			  ForEach ($Member in $s)
			  	{
				  Write-Host $Member.Name + $Member.PrimarySmtpAddress  + $Member.DisplayName
				}
			
			$Recordset.MoveNext()
		} until ($Recordset.EOF -eq $True)
      $RecordSet.Close()
     }
 $connection.Close()
 Remove-PSSession $session
 
   