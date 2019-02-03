<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Copy-ADUser
.SYNOPSIS
    Copies a Active Directory user from another. Including groups and details
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$CopyFormInput                   = New-Object system.Windows.Forms.Form
$CopyFormInput.ClientSize        = '802,728'
$CopyFormInput.text              = "Kopieer Active Directory Gebruiker "
$CopyFormInput.TopMost           = $false

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Gebruiker zoeken:"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(25,22)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 603
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(146,18)
$TextBox1.Font                   = 'Microsoft Sans Serif,10'

$listview                        = New-Object system.Windows.Forms.ListView
$listview.text                   = "listView"
$listview.width                  = 727
$listview.height                 = 124
$listview.location               = New-Object System.Drawing.Point(21,50)

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $false
$TextBox2.width                  = 727
$TextBox2.height                 = 20
$TextBox2.location               = New-Object System.Drawing.Point(21,226)
$TextBox2.Font                   = 'Microsoft Sans Serif,10'

$Groupbox1                       = New-Object system.Windows.Forms.Groupbox
$Groupbox1.height                = 284
$Groupbox1.width                 = 775
$Groupbox1.text                  = "Bestaande gebruiker selecteren"
$Groupbox1.location              = New-Object System.Drawing.Point(12,19)

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Nieuwe gebruiker kopie van:"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(27,199)
$Label2.Font                     = 'Microsoft Sans Serif,10'

$Groupbox_NewUserDetails         = New-Object system.Windows.Forms.Groupbox
$Groupbox_NewUserDetails.height  = 310
$Groupbox_NewUserDetails.width   = 378
$Groupbox_NewUserDetails.text    = "New User Information"
$Groupbox_NewUserDetails.location  = New-Object System.Drawing.Point(11,320)

$Groupbox_NewUserGroupMembership   = New-Object system.Windows.Forms.Groupbox
$Groupbox_NewUserGroupMembership.height  = 100
$Groupbox_NewUserGroupMembership.width  = 378
$Groupbox_NewUserGroupMembership.text  = "Groups"
$Groupbox_NewUserGroupMembership.location  = New-Object System.Drawing.Point(409,319)

$DataGridView1                   = New-Object system.Windows.Forms.DataGridView
$DataGridView1.text              = "User Info"
$DataGridView1.width             = 359
$DataGridView1.height            = 279
$DataGridView1Data = @(@("First Name",""),@("Last Name",""),@("Username",""),@("Business",""),@("Job Title",""))
$DataGridView1.ColumnCount = 2
$DataGridView1.ColumnHeadersVisible = $true
$DataGridView1.Columns[0].Name = "Name"
$DataGridView1.Columns[1].Name = "Value"
foreach ($row in $DataGridView1Data){
          $DataGridView1.Rows.Add($row)
      }
$DataGridView1.Anchor            = 'top,right,bottom,left'
$DataGridView1.location          = New-Object System.Drawing.Point(9,21)

$Groupbox1.controls.AddRange(@($Label1,$TextBox1,$listview,$TextBox2,$Label2))
$CopyFormInput.controls.AddRange(@($Groupbox1,$Groupbox_NewUserDetails,$Groupbox_NewUserGroupMembership))
$Groupbox_NewUserDetails.controls.AddRange(@($DataGridView1))

$listview.Add_SelectedIndexChanged({ UserSearch_OnSelect })
$TextBox1.Add_KeyUp({ tmp })

