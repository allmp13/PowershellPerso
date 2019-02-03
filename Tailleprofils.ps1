Install-Module PSFolderSize
Get-Date | Out-File c:\scripts\OK.txt -append -width 120
foreach ($serverName in (get-content c:\scripts\servers.txt))
{
$serverName
$ping = new-object System.Net.Networkinformation.Ping 
if($ping.send($serverName,"50").status -eq "Success")
{
$serverName | Out-File c:\scripts\OK.txt -append -width 120
$ret="\\" + $serverName + "\c$\users"
Get-FolderSize -Path $ret | Out-File c:\scripts\OK.txt -append -width 120
}
ELSE
{
    $serverName | Out-File c:\scripts\NOK.txt -append -width 120
}
}

