try
{
    Remove-Item $("C:\unfichier.txt") -Force -ErrorAction SilentlyContinue 
    Get-WmiObject -class Win32_ComputerSystem  | Select-Object -ExpandProperty SystemFamily -ErrorAction Stop
    Get-WmiObject Win32_ComputerSystemProduct -ErrorAction SilentlyContinue | Select-Object -ExpandProperty UUID 
}
catch
{
	Write-host $($_.exception.message)
}