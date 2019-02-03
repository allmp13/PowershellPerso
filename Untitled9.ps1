$dataSource = "srvmssql-bi"
$user = "PROG_FIN_GSI_ADMIN"
$pwd = "201509151546A"
$database = "PROG_FIN_GSI"
$connectionString = "Server=$dataSource;uid=$user; pwd=$pwd;Database=$database;Integrated Security=False;"
$SqlQuery ="SELECT Programmes.ID, Programmes.[Libelle Programme], Programmes.LC, Programmes.Operation, Programmes.[Operation Nature], Programmes.Exercice FROM Programmes WHERE (((Programmes.Exercice)='2019')) ORDER BY Programmes.[Libelle Programme]"



$dt = New-Object System.Data.DataTable 
$da = New-Object System.Data.SqlClient.SqlDataAdapter -ArgumentList  $SqlQuery,$connectionString
$da.fill($dt)| Out-Null
#$dt.Rows[120].'Libelle Programme'

<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '600,250'
$Form.text                       = "Saisie Engagement / Programme"
$Form.TopMost                    = $true

$ComboBox1                       = New-Object system.Windows.Forms.ComboBox
$ComboBox1.text                  = "comboBox"
$ComboBox1.width                 = 500
$ComboBox1.height                = 20
$ComboBox1.location              = New-Object System.Drawing.Point(60,74)
$ComboBox1.Font                  = 'Microsoft Sans Serif,10'

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Valider"
$Button1.width                   = 60
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(290,202)
$Button1.Font                    = 'Microsoft Sans Serif,10'

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 170
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(60,119)
$TextBox1.Font                   = 'Microsoft Sans Serif,10'

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Sortir"
$Button2.width                   = 60
$Button2.height                  = 30
$Button2.location                = New-Object System.Drawing.Point(376,202)
$Button2.Font                    = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($ComboBox1,$Button1,$TextBox1,$Button2))

$ComboBox1.ValueMember = "ID"
$ComboBox1.DisplayMember = "Libelle Programme" 
$ComboBox1.Datasource = $dt

$Button1.Add_Click({Enregistrer})
$Button2.Add_Click({ closeForm })


#Write your logic code here

[void]$Form.ShowDialog()

function Enregistrer(){
    if ($TextBox1.Text) #Not NULL
    {
        Write-Host $Combobox1.SelectedItem["Libelle Programme"]  " --> " $Combobox1.SelectedItem["ID"]
        $Sqlinsert="INSERT INTO Executions (IDPROG, NUMENGAGEMENT, EXERCICE) VALUES ('"+$Combobox1.SelectedItem["ID"]+"', '"+$TextBox1.Text+"', '2019')"
        $daa = New-Object System.Data.SqlClient.SqlDataAdapter -ArgumentList  $Sqlinsert,$connectionString
        $dtt = New-Object System.Data.DataTable
        $daa.fill($dtt)| Out-Null
        $TextBox1.Clear()
    }
}

function closeForm(){
    $Form.close()
}