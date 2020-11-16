
((Get-Item -Path "C:\Users\cpages").LastAccessTime)

if (((Get-Item -Path "C:\Users\cpages").LastWriteTime) -lt (Get-Date).AddMonths(-2))
{ $perime = $true}
else {
    $perime = $false
}
$perime