Get-WmiObject -Class Win32_UserProfile -ComputerName B19U352  |      
Where-Object { $_.LocalPath -Like "*ure" } |

ForEach-Object {
    $_
    Write-Host $_.LocalPath
    Write-Host $_.LastUseTime
    $jmd = ([WMI] '').ConvertToDateTime($_.LastUseTime)
    Write-Host $jmd
    if (([WMI] '').ConvertToDateTime($_.LastUseTime) -lt (Get-Date).Addmonths(-6))
    { Write-Host Supérieur à 1 jour }
    else {
        Write-Host OK!
    }
}