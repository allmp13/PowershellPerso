# psexec \\b14u038 cmd /c "powershell -noprofile -executionpolicy bypass -file C:\temp\LocalProfil.ps1"


Start-Transcript -Path "C:\temp\log.txt"
$start=(get-date)
Import-Module c:\temp\PSFolderSize
#$ret=Read-Host "Please enter a value"
Get-FolderSize -BasePath c:\users | Out-File c:\temp\zuzu1.txt -append -width 120
#$ret=Read-Host "Please enter a value"
$end=(get-date)
Out-File c:\temp\zuzu1.txt -inputobject "debut $start fin $end Duree du script $(($end-$start).totalseconds) Secondes"  -append -width 120 
Stop-Transcript