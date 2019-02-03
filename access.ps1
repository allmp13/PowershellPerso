Function Check-Path($Db)
{
 If(!(Test-Path -path (Split-Path -path $Db -parent)))
   { 
     Throw "$(Split-Path -path $Db -parent) Does not Exist" 
   }
  ELSE
  { 
   If(!(Test-Path -Path $Db))
     {
      Throw "$db does not exist"
     }
  }
} #End Check-Path

Function Get-Bios
{
 Get-WmiObject -Class Win32_Bios
} #End Get-Bios

Function Get-Video
{
 Get-WmiObject -Class Win32_VideoController
} #End Get-Video

Function Connect-Database($Db, $Tables)
{
  $OpenStatic = 3
  $LockOptimistic = 3
  $connection = New-Object -ComObject ADODB.Connection
  $connection.Open("Provider = Microsoft.ACE.OLEDB.12.0;Data Source=$Db" )
  Update-Records($Tables)
} #End Connect-DataBase

Function Update-Records($Tables)
{
  $RecordSet = new-object -ComObject ADODB.Recordset
   ForEach($Table in $Tables)
     {
      $Query = "Select * from $Table"
      $RecordSet.Open($Query, $Connection, $OpenStatic, $LockOptimistic)
      Invoke-Expression "Update-$Table"
      $RecordSet.Close()
     }
   $connection.Close()
} #End Update-Records

Function Update-matable
{
 "Updating Bios"
 $BiosInfo = Get-Bios

 $RecordSet.AddNew()
 $RecordSet.Fields.Item("nom") = $BiosInfo.SMBIOSBIOSVersion
 $RecordSet.Fields.Item("prenom") = $BiosInfo.Version
 $RecordSet.Update()
} #End Update-Bios

Function Update-Video
{
 "Updating video"
 $VideoInformation = Get-Video
 Foreach($VideoInfo in $VideoInformation)
  { 
   $RecordSet.AddNew()
   $RecordSet.Fields.Item("nom") = $VideoInfo.AdapterCompatibility
   $RecordSet.Fields.Item("prenom") = $VideoInfo.AdapterDACType
   $RecordSet.Update()
  }
} #End Update-Video

# *** Entry Point to Script ***
"Debut"
#$Db = "C:\temp\jmd.accdb"
$Db = "jmd.accdb"
$Tables = "matable"
#Check-Path -db $Db
Connect-DataBase -db $Db -tables $Tables
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
 
Write-Host "Current script directory is $ScriptDir"
"Fin"