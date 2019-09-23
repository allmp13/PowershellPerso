$dataSource = "srvmssql-p"
$user = "PAM"
$pwd = "DELPHES"
$database = "PAM3DATA"
$connectionString = "Server=$dataSource;uid=$user; pwd=$pwd;Database=$database;Integrated Security=False;"
Install-Module PSFolderSize -Force
Get-Date | Out-File ..\OK.txt -append -width 120
foreach ($serverName in (get-content ..\servers.txt)) {
    $serverName = $serverName.Substring(0, 7)
    #$serverName
    #$Sqlinsert = "select o.odDateScan from otObject o where o.otInternalNumber='" + $serverName + "'"
    $Sqlinsert ="SELECT otObject.odDateScan, [otobject_1].[otModel]+'  \ '+[otobject_2].[otModel]+'  \ '+[otobject_3].[otModel] AS Nom
    FROM (((otObject INNER JOIN otObjectRelation ON otObject.oidObject = otObjectRelation.oidObject) INNER JOIN otObject AS otObject_1 ON otObjectRelation.oidObjectRelation = otObject_1.oidObject) INNER JOIN (otObjectRelation AS otObjectRelation_1 INNER JOIN otObject AS otObject_2 ON otObjectRelation_1.oidObjectRelation = otObject_2.oidObject) ON otObject.oidObject = otObjectRelation_1.oidObject) INNER JOIN (otObjectRelation AS otObjectRelation_2 INNER JOIN otObject AS otObject_3 ON otObjectRelation_2.oidObjectRelation = otObject_3.oidObject) ON otObject.oidObject = otObjectRelation_2.oidObject
    WHERE (((otObject.otInternalNumber)='" + $serverName + "') AND ((otObjectRelation.oidObjectRelationType)=201) AND ((otObjectRelation_1.oidObjectRelationType)=302) AND ((otObjectRelation_2.oidObjectRelationType)=304))"
    $SqlConnection = new-object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = $connectionString
    $SqlCommand = $SqlConnection.CreateCommand()
    $SqlCommand.CommandText = $Sqlinsert
    $DataAdapter = new-object System.Data.SqlClient.SqlDataAdapter $SqlCommand
    $dataset = new-object System.Data.Dataset
    $DataAdapter.Fill($dataset) | Out-Null
    $dateinventaire = $dataset.Tables[0].Rows[0].Item(0)
    $messageinventaire = $serverName + " \ " + $dataset.Tables[0].Rows[0].Item(1) + " --> Dernier inventaire le: " + $dateinventaire.tostring("dd-MM-yyyy") + " Testé le : " + (Get-Date).tostring("dd-MM-yyyy")
    if ($dateinventaire -lt (Get-Date).AddMonths(-4)) { $messageinventaire = $messageinventaire + ' Non utilisé depuis plus de 4 Mois!!!!' } else { $messageinventaire = $messageinventaire + " OK!" } 
    $messageinventaire
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

