Clear-Host
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
	  Write-Host $fin
      $RecordSet.movefirst()
	  do 
	  	{
			for($i=0; $i -lt $fin; $i++){$Recordset.Fields.Item($i).Value}
			 
			$Recordset.MoveNext()
		} until ($Recordset.EOF -eq $True)
      $RecordSet.Close()
     }
 $connection.Close()