#requires -RunAsAdministrator
function DisplayInBytes($num) 
	{
		$suffix = "B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"
		$index = 0
		while ($num -gt 1kb) 
		{
			$num = $num / 1kb
			$index++
		} 
	
		"{0:N2} {1}" -f $num, $suffix[$index]
	}
#Full path and filename of the file to write the output to
$File = "c:\ProfileSizeReport.csv"
#Exclude these accounts from reporting
$Exclusions = ("Administrateur","Admin", "Default", "Public")
#Get the list of profiles
$Profiles = Get-ChildItem -Path "\\b16u016\c$\Users" | Where-Object { $_ -notin $Exclusions }
#Create the object array
$AllProfiles = @()
#Create the custom object
foreach ($Profile in $Profiles) {
	$object = New-Object -TypeName System.Management.Automation.PSObject
	#Get the size of the Documents and Desktop combined and round with no decimal places
	$folderSizes= (Get-ChildItem -path $Profile.FullName -Recurse | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum
	#$folderSizes = "{0} MB" -f ($folderSizes / 1MB)
	Write-Host $folderSizes
	$folderSizes=DisplayInBytes($folderSizes)
	Write-Host $folderSizes
	$object | Add-Member -MemberType NoteProperty -Name ComputerName -Value $env:COMPUTERNAME.ToUpper()
	$object | Add-Member -MemberType NoteProperty -Name Profile -Value $Profile
	$Object | Add-Member -MemberType NoteProperty -Name Size -Value $FolderSizes
	$AllProfiles += $object
}
#Create the formatted entry to write to the file
[string]$Output = $null
foreach ($Entry in $AllProfiles) {
	[string]$Output += $Entry.ComputerName + ';' + $Entry.Profile + ';' + $Entry.Size + [char]13
}
#Remove the last line break
$Output = $Output.Substring(0,$Output.Length-1)
#Write the output to the specified CSV file. If the file is opened by another machine, continue trying to open until successful
Do {
	Try {
		$Output | Out-File -FilePath $File -Encoding UTF8 -Append -Force
		$Success = $true
	} Catch {
		$Success = $false
	}
	Write-Host $Success
} while ($Success -eq $false)
Write-Host $Success
