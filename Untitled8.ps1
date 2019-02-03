[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Windows.Forms.Application]::EnableVisualStyles()

#create a form
$form = New-Object System.Windows.Forms.Form

#create a datatable to bind to our combobox
$datatable = New-Object system.Data.DataTable
		
#Define Columns
$col1 = New-Object system.Data.DataColumn "Value",([string])
$col2 = New-Object system.Data.DataColumn "Text",([string])

#add columns to datatable
$datatable.columns.add($col1)
$datatable.columns.add($col2)
		
#Create a row
$datarow1 = $datatable.NewRow()

#Enter data in the row
$datarow1.Value = "Example Value 1"
$datarow1.Text = "Example Text 1"

#Add the row to the datatable
$datatable.Rows.Add($datarow1)

#Create another row
$datarow2 = $datatable.NewRow()

#Enter data in the row
$datarow2.Value = "Example Value 2"
$datarow2.Text = "Example Text 2"

#Add the row to the datatable
$datatable.Rows.Add($datarow2)

#create a combobox
$combobox = New-Object System.Windows.Forms.ComboBox		
$combobox.Add_SelectedIndexChanged({
		#output the selected value and text
		write-host $combobox.SelectedItem["Value"] $combobox.SelectedItem["Text"]
})

#clear combo before we bind it
$combobox.Items.Clear()

#bind combobox to datatable
$combobox.ValueMember = "Value"
$combobox.DisplayMember = "Text"
$combobox.Datasource = $datatable

#add combobox to form
$form.Controls.Add($combobox)	

#show form
[void]$form.showdialog()