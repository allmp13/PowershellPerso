$SourceFile = "C:\Temp\File.txt"
$DestinationFile = "C:\Temp\NonexistentDirectory\File.txt"
Write-Host Test-Path $DestinationFile

If (Test-Path $DestinationFile) {
    $i = 0
    While (Test-Path $DestinationFile) {
	Write-Host Test-Path $DestinationFile
        $i += 1
        $DestinationFile = "C:\Temp\NonexistentDirectory\File$i.txt"

    }
} Else {
    New-Item -ItemType File -Path $DestinationFile -Force
}

Copy-Item -Path $SourceFile -Destination $DestinationFile -Force
Get-DistributionGroup
