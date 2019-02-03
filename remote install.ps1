$file = '\\srvbur\public\ocs\OCS-NG-Windows-Agent-Setup.exe'
$computerName = 'B11U259'

Copy-Item -Path $file -Destination "\\$computername\c$\temp\OCS-NG-Windows-Agent-Setup.exe"
Invoke-Command -ComputerName $computerName -ScriptBlock {
    Start-Process c:\temp\OCS-NG-Windows-Agent-Setup.exe -ArgumentList '/S /NO_SYSTRAY /NOSPLASH /SERVER=http://srvocs.sdis13.fr/ocsinventory /NOW' -Wait
}
