$key = get-item "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\"
$val = get-itemProperty $key.PSPath |Select-Object ProductName
$val
[System.Net.Dns]::Resolve('www.google.fr').AddressList.IPaddresstostring
Get-WmiObject Win32_Computersystem
Get-WmiObject Win32_Bios
Get-WmiObject Win32_LogicalDisk
Get-WmiObject Win32_MemoryDevice
Get-WmiObject –class Win32_NetworkAdapterConfiguration