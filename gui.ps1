function jmd 
{
<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    JM
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$FORM                            = New-Object system.Windows.Forms.Form
$FORM.ClientSize                 = '400,400'
$FORM.text                       = "Mon GUI"
$FORM.TopMost                    = $true

$OK                              = New-Object system.Windows.Forms.Button
$OK.text                         = "Ok"
$OK.width                        = 60
$OK.height                       = 30
$OK.location                     = New-Object System.Drawing.Point(56,295)
$OK.Font                         = 'Microsoft Sans Serif,10'

$Annuler                         = New-Object system.Windows.Forms.Button
$Annuler.text                    = "Annuler"
$Annuler.width                   = 60
$Annuler.height                  = 30
$Annuler.location                = New-Object System.Drawing.Point(132,295)
$Annuler.Font                    = 'Microsoft Sans Serif,10'

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 100
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(88,63)
$TextBox1.Font                   = 'Microsoft Sans Serif,10'
$TextBox1.name = "toto"

$FORM.controls.AddRange(@($OK,$Annuler,$TextBox1))

$OK.Add_Click({ closeForm })
$Annuler.Add_Click({ annuler })
function closeForm(){
    $Form.close()
    Write-Host $TextBox1.Text
}
function annuler(){
    $TextBox1.Text="J'annule"
}
 [void]$Form.ShowDialog()

 }


Write-Output "Début"
jmd