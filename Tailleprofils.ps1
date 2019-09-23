$dataSource = "srvmssql-p"
$user = "PAM"
$pwd = "DELPHES"
$database = "PAM3DATA"
$connectionString = "Server=$dataSource;uid=$user; pwd=$pwd;Database=$database;Integrated Security=False;"
Install-Module PSFolderSize -Force
Get-Date | Out-File ..\OK.txt -append -width 120
foreach ($serverName in (get-content ..\servers.txt)) {
    $serverName = $serverName.Substring(0, 7)
    $serverName
    $Sqlinsert = "select o.odDateScan from otObject o where o.otInternalNumber='" + $serverName + "'"
    $SqlConnection = new-object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = $connectionString
    $SqlCommand = $SqlConnection.CreateCommand()
    $SqlCommand.CommandText = $Sqlinsert
    $DataAdapter = new-object System.Data.SqlClient.SqlDataAdapter $SqlCommand
    $dataset = new-object System.Data.Dataset
    $DataAdapter.Fill($dataset) | Out-Null
    $dateinventaire = $dataset.Tables[0].Rows[0].Item(0)
    $messageinventaire = $serverName + " --> Dernier inventaire le: " + $dateinventaire.tostring("dd-MM-yyyy") + " Testé le : " + (Get-Date).tostring("dd-MM-yyyy")
    if ($dateinventaire -lt (Get-Date).AddMonths(-4)) { $messageinventaire = $messageinventaire + ' Non utilisé depuis plus de 4 Mois!!!!' } else { $messageinventaire = $messageinventaire + " OK!" } 

    #"Dernier inventaire le: " + $dataset.Tables[0].Rows[0].Item(0) 

    if ((Test-Connection -computername $serverName -Count 1 -quie) -eq "Success") {
        $messageinventaire | Out-File ..\OK.txt -append -width 120
        $ret = "\\" + $serverName + "\c$\users"
        Get-FolderSize -Path $ret | Out-File ..\OK.txt -append -width 120
    }
    else {
        $messageinventaire | Out-File ..\NOK.txt -append -width 120
    }
}

