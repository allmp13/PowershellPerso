$XLSDoc = "C:\temp\Liste des unités centrales 2020.xlsx"
$SheetName = "Firstname"
$Excel = New-Object -ComObject "Excel.Application"

$Workbook = $Excel.workbooks.open($XLSDoc)
$Sheet = $Workbook.Worksheets.Item(1)
$Excel.Visible = $false

$RowCount = $Sheet.UsedRange.Rows.Count
Write-Host "Il y a $RowCount lignes"

for ($i=2; $i -le $RowCount; $i++){
$lastname = $Sheet.Cells.Item($i,1).Text
$Firstname = $Sheet.Cells.Item($i,2).Text
$username = $Sheet.Cells.Item($i,3).Text
$groupe = $Sheet.Cells.Item($i,4).Text
$Sheet.Cells.Item($i,22)="Ligne N°$i"
$test = $Sheet.Cells.Item($i,22).Text


Write-host $firstname $lastname $Username $groupe $test
}
$Workbook.Save()
$Workbook.Close()
$Excel.quit()