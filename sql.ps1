srvmssql-bi
$dataSource = "srvmssql-bi"
$user = "PROG_FIN_GSI_ADMIN"
$pwd = "201509151546A"
$database = "PROG_FIN_GSI"
$connectionString = "Server=$dataSource;uid=$user; pwd=$pwd;Database=$database;Integrated Security=False;"

##$query = "SELECT * FROM Executions"

$query ="SELECT Programmes.ID, Programmes.[Libelle Programme], Programmes.LC, Programmes.Operation, Programmes.[Operation Nature], Programmes.Exercice FROM Programmes WHERE (((Programmes.Exercice)='2019')) ORDER BY Programmes.[Libelle Programme]"

$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString
#$connection.ConnectionString = "Server=$dataSource;Database=$database;Integrated Security=True;"
$connection.Open()
$command = $connection.CreateCommand()
$command.CommandText = $query

$result = $command.ExecuteReader()


$table = new-object “System.Data.DataTable”
$table.Load($result)

$table



$connection.Close()



