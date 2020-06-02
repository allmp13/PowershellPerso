########################################################
#Jean-Marc DAVID le 01/02/2019
# Pour le compiler en exe utiliser PS2EXE-GUI
# Pour créer l'interface Graphique Utiliser POSHGUI.com


add-Type -AssemblyName system.data
$anneeexercice='2019'
$dataSource = "srvmssql-bi"
$user = "PROG_FIN_GSI_ADMIN"
$pwd = "201509151546A"
$database = "PROG_FIN_GSI"
$connectionString = "Server=$dataSource;uid=$user; pwd=$pwd;Database=$database;Integrated Security=False;"
#$SqlQuery ="SELECT Programmes.ID, Programmes.[Libelle Programme], Programmes.LC, Programmes.Operation, Programmes.[Operation Nature], Programmes.Exercice FROM Programmes WHERE (((Programmes.Exercice)='2019')) ORDER BY Programmes.[Libelle Programme]"
$SqlQuery ="SELECT Programmes.ID, Programmes.[Libelle Programme] + ' , LC: ' + Programmes.LC AS [Libelle Programme] FROM Programmes  WHERE (((Programmes.Exercice)='"+$anneeexercice+"')) ORDER BY Programmes.[Libelle Programme]"

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
$Form.ClientSize                 = '800,204'
$Form.text                       = "Saisie Engagement / Programme " + $anneeexercice
$Form.TopMost                    = $true

$ComboBox1                       = New-Object system.Windows.Forms.ComboBox
$ComboBox1.text                  = "comboBox"
$ComboBox1.width                 = 780
$ComboBox1.height                = 20
$ComboBox1.location              = New-Object System.Drawing.Point(10,74)
$ComboBox1.Font                  = 'Microsoft Sans Serif,10'

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Valider"
$Button1.width                   = 60
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(348,152)
$Button1.Font                    = 'Microsoft Sans Serif,10'

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 63
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(10,132)
$TextBox1.Font                   = 'Microsoft Sans Serif,10'

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Sortir"
$Button2.width                   = 60
$Button2.height                  = 30
$Button2.location                = New-Object System.Drawing.Point(434,152)
$Button2.Font                    = 'Microsoft Sans Serif,10'

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Libellé Programme"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(10,51)
$Label1.Font                     = 'Microsoft Sans Serif,10'
$Label1.ForeColor                = "#d0021b"

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Numéro Engagement Astre"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(10,110)
$Label2.Font                     = 'Microsoft Sans Serif,10'
$Label2.ForeColor                = "#d0021b"

$Form.controls.AddRange(@($ComboBox1,$Button1,$TextBox1,$Button2,$Label1,$Label2))
$ComboBox1.ValueMember = "ID"
$ComboBox1.DisplayMember = "Libelle Programme" 
$ComboBox1.Datasource = $dt

$Button1.Add_Click({Enregistrer})
$Button2.Add_Click({ closeForm })


#Write your logic code here

function Enregistrer(){
    if (($TextBox1.Text) -and ($TextBox1.Text.Length -lt 8)) #Not NULL
    {
        #Write-Host $Combobox1.SelectedItem["Libelle Programme"]  " --> " $Combobox1.SelectedItem["ID"]
        $Sqlinsert="INSERT INTO Executions (IDPROG, NUMENGAGEMENT, EXERCICE) VALUES ('"+$Combobox1.SelectedItem["ID"]+"', '"+$TextBox1.Text.ToUpper()+"', '"+$anneeexercice+"')"
        $daa = New-Object System.Data.SqlClient.SqlDataAdapter -ArgumentList  $Sqlinsert,$connectionString
        $dtt = New-Object System.Data.DataTable
        $daa.fill($dtt)| Out-Null
        $TextBox1.Clear()
    }
}

function closeForm(){
    $Form.close()
}

$showWindowAsync = Add-Type –memberDefinition @” 
[DllImport("user32.dll")] 
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow); 
“@ -name “Win32ShowWindowAsync” -namespace Win32Functions –passThru

function Show-PowerShell() { 
     [void]$showWindowAsync::ShowWindowAsync((Get-Process –id $pid).MainWindowHandle, 10) 
}

function Hide-PowerShell() { 
    [void]$showWindowAsync::ShowWindowAsync((Get-Process –id $pid).MainWindowHandle, 2) 
}

Hide-PowerShell

[void]$Form.ShowDialog()

