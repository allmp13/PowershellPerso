#The math works only with seconds
$pauseSecs = 1
$MaxWaitSecs =15

for ($i=0; $i -lt $MaxWaitSecs; $i+=$pauseSecs) {
 [int]$pct = ($i/$MaxWaitSecs)*100 
 $Params = @{
     Activity = "Please wait"
     PercentComplet = $pct
     CurrentOperation  = "$i seconds elapsed, $pct% of maximum $MaxWaitSecs seconds wait time" 
     Status = "Pausing"
 }
 Write-Progress @params
 Start-Sleep -seconds $pauseSecs 
}