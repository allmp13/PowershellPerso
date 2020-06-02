$dataSource = "srvmssql-p"
$user = "PAM"
$pwd = "DELPHES"
$database = "PAM3DATA"
$connectionString = "Server=$dataSource;uid=$user; pwd=$pwd;Database=$database;Integrated Security=False;"

$XLSDoc = "C:\temp\Liste des unités centrales 2020.xlsx"
$SheetName = "Firstname"
$Excel = New-Object -ComObject "Excel.Application"

$Workbook = $Excel.workbooks.open($XLSDoc)
$Sheet = $Workbook.Worksheets.Item(1)
$Excel.Visible = $false

$RowCount = $Sheet.UsedRange.Rows.Count
Write-Host "Il y a $RowCount lignes"

for ($i=2; $i -le $RowCount; $i++){
$serverName = $Sheet.Cells.Item($i,1).Text
$Firstname = $Sheet.Cells.Item($i,2).Text
$username = $Sheet.Cells.Item($i,3).Text
$groupe = $Sheet.Cells.Item($i,4).Text


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
$Sheet.Cells.Item($i,22)=$messageinventaire

Write-host $serverName
}
$Workbook.Save()
$Workbook.Close()
$Excel.quit()



