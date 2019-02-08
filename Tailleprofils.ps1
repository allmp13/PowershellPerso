$dataSource = "srvmssql-p"
$user = "PAM"
$pwd = "DELPHES"
$database = "PAM3DATA"
$connectionString = "Server=$dataSource;uid=$user; pwd=$pwd;Database=$database;Integrated Security=False;"
Install-Module PSFolderSize
Get-Date | Out-File c:\scripts\OK.txt -append -width 120
foreach ($serverName in (get-content c:\scripts\servers.txt))
{
$serverName=$serverName.Substring(0,7)
$serverName
$Sqlinsert="select o.odDateScan from otObject o where o.otInternalNumber='" + $serverName + "'"
$SqlConnection = new-object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = $connectionString
$SqlCommand = $SqlConnection.CreateCommand()
$SqlCommand.CommandText = $Sqlinsert
$DataAdapter = new-object System.Data.SqlClient.SqlDataAdapter $SqlCommand
$dataset = new-object System.Data.Dataset
$DataAdapter.Fill($dataset)| Out-Null
$dateinventaire=$dataset.Tables[0].Rows[0].Item(0)
$messageinventaire=$serverName + " --> Dernier inventaire le: " + $dateinventaire.tostring("dd-MM-yyyy") + " Testé le : " + (Get-Date).tostring("dd-MM-yyyy")
if ($dateinventaire -lt (Get-Date).AddMonths(-4)) {$messageinventaire=$messageinventaire + ' Non utilisé depuis plus de 4 Mois!!!!'} else {$messageinventaire=$messageinventaire + " OK!"} 

#"Dernier inventaire le: " + $dataset.Tables[0].Rows[0].Item(0) 


$ping = new-object System.Net.Networkinformation.Ping 
if($ping.send($serverName,"50").status -eq "Success")
{
$messageinventaire | Out-File c:\scripts\OK.txt -append -width 120
$ret="\\" + $serverName + "\c$\users"
Get-FolderSize -Path $ret | Out-File c:\scripts\OK.txt -append -width 120
}
ELSE
{
    $messageinventaire  | Out-File c:\scripts\NOK.txt -append -width 120
}
}

