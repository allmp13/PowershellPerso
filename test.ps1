$Profiles = Get-ChildItem -Path $env:SystemDrive"\Users"
Write-Host $env:SystemDrive"\Users"
$folderSize=(Get-ChildItem ($Profile.FullName + '\Documents'), ($Profile.FullName + '\Desktop') -Recurse | Measure-Object -Property Length -Sum -ErrorAction Stop).sum
$folderSizeInKB = "{0} KB" -f ($folderSize / 1KB)
$folderSizeInMB = "{0} MB" -f ($folderSize / 1MB)
$folderSizeInGB = "{0} GB" -f ($folderSize / 1GB)
Write-Host  $folderSize
write-Host $folderSizeInKB
Write-Host $folderSizeInMB
Write-Host $folderSizeInGB

$FolderSizes = "{0:N2}" -f ((Get-ChildItem ($Profile.FullName + '\Documents'), ($Profile.FullName + '\Desktop') -Recurse | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum)
Write-Host  $folderSizes
$FolderSizes = [System.Math]::round($folderSizes,2)