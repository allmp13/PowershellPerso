<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '800,500'
$Form.text                       = "Form"
$Form.TopMost                    = $false

$ComboBox1                       = New-Object system.Windows.Forms.ComboBox
$ComboBox1.text                  = "comboBox"
$ComboBox1.width                 = 700
$ComboBox1.height                = 20
$ComboBox1.location              = New-Object System.Drawing.Point(50,94)
$ComboBox1.Font                  = 'Microsoft Sans Serif,10'

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "button"
$Button1.width                   = 60
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(244,250)
$Button1.Font                    = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($ComboBox1,$Button1))

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

#while ($result.Read()) { $result["Libelle Programme"] + ", " + $result["LC"]}

while ($result.Read()) { 
    $ComboBox1.Items.Add($result["Libelle Programme"]+", LC: " + $result["LC"])
}
$Combobox1.SelectedIndex=0


$Button1.Add_Click({Write-Host $Combobox1.Text})


<#$table = new-object “System.Data.DataTable”
$table.Load($result)

$table#>



$connection.Close()



#Write your logic code here

[void]$Form.ShowDialog()

FUNCTION TOTO(){
Write-Host $Combobox1.DisplayMember
}
