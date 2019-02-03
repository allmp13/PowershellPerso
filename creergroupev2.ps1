Clear-Host
$password = Cat c:\scripts\MonFichier.pwd | ConvertTo-SecureString
$user = "ddsis-13\jmdavid"
$credential = New-Object System.Management.Automation.PSCredential($user,$password)
$session = New-PSSession -ConfigurationName Microsoft.Exchange –ConnectionUri http://srvcas1/powershell -Credential $credential
$output=Import-PSSession $session -AllowClobber -DisableNameChecking  
Import-Module ActiveDirectory



$Db = "c:\scripts\jmd.accdb"
#$Table = "reference2"
$Table = "listecs"
$OpenStatic = 3
$LockOptimistic = 3
$connection = New-Object -ComObject ADODB.Connection
$connection.Open("Provider = Microsoft.ACE.OLEDB.12.0;Data Source=$Db" )
$RecordSet = new-object -ComObject ADODB.Recordset
      $Query = "Select * from $Table"
      $RecordSet.Open($Query, $Connection, $OpenStatic, $LockOptimistic)
	  $fin = $RecordSet.Fields.Count
	  #Write-Host "$fin Champ(s)"
      $RecordSet.movefirst()
	  do 
	  	{
			<#$s=Get-DistributionGroupMember -Identity $Recordset.Fields.Item(3).value | select displayname, Name, PrimarySmtpAddress
			$group=Get-DistributionGroup -Identity $Recordset.Fields.Item(3).Value | select displayname
			Write-Host "----------------------------------------------------------------------------------------"
			Write-Host $group.DisplayName
			Write-Host "----------------------------------------------------------------------------------------"
			  ForEach ($Member in $s)
			  	{
				  Write-Host $Member.Name #$Member.PrimarySmtpAddress $Member.DisplayName
				}
			Write-Host#>
			
			
	

$Udept = $Recordset.Fields.Item(0).value            	# Trigramme du département
$dept=$Udept.tolower()


#Write-Host (Get-ADOrganizationalUnit -Filter 'Name -like $dept').distinguishedName
$correspondantinformatique = $dept + "inf" 	# Alias du groupe de distribution informatique du département
$correspondantpatrimoine = $dept + "pat" 	# Alias du groupe de distribution patrimoine du département
#Write-Host $correspondantinformatique
#Write-Host $correspondantpatrimoine

$chefdecentre = $dept + "chf"
$adjoint = $dept + "adj"
#Write-Host $chefdecentre
#Write-Host $adjoint


Write-Host $Udept 
#Write-Host $dept 



# Correspondant Informatique

# Groupe de distribution Correspondant Informatique

if (((Get-DistributionGroup -Identity $correspondantinformatique -ErrorAction 'SilentlyContinue').IsValid) -eq $true)
{
	#Write-Host "Groupe de Distribution $correspondantinformatique OK !"
	#Write-Host (Get-DistributionGroup $correspondantinformatique).OrganizationalUnit
	Write-Host (Get-DistributionGroup $correspondantinformatique | select name) (Get-DistributionGroupMember $correspondantinformatique).Count (Get-DistributionGroupMember $correspondantinformatique| select name)
}
else
{
	Write-Host "Groupe de Distribution $correspondantinformatique Inconnu !"
}

# Membres du groupe de distribution Correspondant Informatique


if (((Get-DistributionGroup -Identity $chefdecentre -ErrorAction 'SilentlyContinue').IsValid) -eq $true)
{

	if ((Get-DistributionGroupMember $correspondantinformatique | select -expand distinguishedname) -contains (Get-DistributionGroup $chefdecentre).distinguishedname)
	{
		#Write-Host "membre $chefdecentre dans groupe $correspondantinformatique"
		Write-Host (Get-DistributionGroup $chefdecentre | select name) (Get-DistributionGroupMember $chefdecentre).Count (Get-DistributionGroupMember $chefdecentre| select name)

	}
	else
	{
		Write-Host "ATTENTION!!!! $chefdecentre n est pas membre du groupe $correspondantinformatique"
		Add-DistributionGroupMember -Identity $correspondantinformatique -Member $chefdecentre
	}	
}

if (((Get-DistributionGroup -Identity $adjoint -ErrorAction 'SilentlyContinue').IsValid) -eq $true)
{

	if ((Get-DistributionGroupMember $correspondantinformatique | select -expand distinguishedname) -contains (Get-DistributionGroup $adjoint).distinguishedname)
	{
		#Write-Host "membre $adjoint dans groupe $correspondantinformatique"
		Write-Host (Get-DistributionGroup $adjoint | select name) (Get-DistributionGroupMember $adjoint).Count (Get-DistributionGroupMember $adjoint| select name)
		
	}
	else
	{
		Write-Host "ATTENTION!!!! $adjoint n est pas membre du groupe $correspondantinformatique"
		Add-DistributionGroupMember -Identity $correspondantinformatique -Member $adjoint

	}
}



# Correspondant Patrimoine

if (((Get-DistributionGroup -Identity $correspondantpatrimoine -ErrorAction 'SilentlyContinue').IsValid) -eq $true)
{
	#Write-Host "Groupe de Distribution $correspondantpatrimoine OK !"
	Write-Host (Get-DistributionGroup $correspondantpatrimoine | select name) (Get-DistributionGroupMember $correspondantpatrimoine).Count (Get-DistributionGroupMember $correspondantpatrimoine| select name)
	#Remove-DistributionGroupMember -Identity  $correspondantpatrimoine -Member $correspondantinformatique  -Confirm:$false
}
else
{
	Write-Host "Groupe de Distribution $correspondantpatrimoine Inconnu !"
	<#
	New-DistributionGroup -OrganizationalUnit  (Get-DistributionGroup $correspondantinformatique).OrganizationalUnit -Alias $correspondantpatrimoine -Name "Correspondant Patrimoine $Udept" 
    Set-DistributionGroup -Identity $correspondantpatrimoine -RequireSenderAuthenticationEnabled $False
	Add-DistributionGroupMember -Identity $correspondantpatrimoine -Member $chefdecentre
	Add-DistributionGroupMember -Identity $correspondantpatrimoine -Member $adjoint
#>
}

if (((Get-DistributionGroup -Identity $chefdecentre -ErrorAction 'SilentlyContinue').IsValid) -eq $true)
{

	if ((Get-DistributionGroupMember $correspondantpatrimoine | select -expand distinguishedname) -contains (Get-DistributionGroup $chefdecentre).distinguishedname)
	{
		#Write-Host "membre $chefdecentre dans groupe $correspondantpatrimoine"
	}
	else
	{
		Write-Host "ATTENTION!!!! $chefdecentre n est pas membre du groupe $correspondantpatrimoine"
	}	
}
else
{
	Write-Host "ATTENTION!!!! $chefdecentre n existe pas"
}

if (((Get-DistributionGroup -Identity $adjoint -ErrorAction 'SilentlyContinue').IsValid) -eq $true)
{

	if ((Get-DistributionGroupMember $correspondantpatrimoine | select -expand distinguishedname) -contains (Get-DistributionGroup $adjoint).distinguishedname)
	{
		#Write-Host "membre $adjoint dans groupe $correspondantpatrimoine"
	}
	else
	{
		Write-Host "ATTENTION!!!! $adjoint n est pas membre du groupe $correspondantpatrimoine"

	}
}
else
{
	Write-Host "ATTENTION!!!! $adjoint n existe pas"
}



			
			$Recordset.MoveNext()
		} until ($Recordset.EOF -eq $True)
      $RecordSet.Close()
 $connection.Close()


<#

$dept = "AIX".ToLower()             	# Trigramme du département
$correspondantinformatique = $dept + "inf" 	# Alias du groupe de distribution informatique du département
$correspondantpatrimoine = $dept + "pat" 	# Alias du groupe de distribution patrimoine du département

Write-Host $dept 
Write-Host $correspondantinformatique
Write-Host $correspondantpatrimoine


# Correspondant Informatique

# Groupe de distribution Correspondant Informatique

if (((Get-DistributionGroup -Identity $correspondantinformatique -ErrorAction 'SilentlyContinue').IsValid) -eq $true)
{
	Write-Host "Groupe de Distribution $correspondantinformatique OK !"
	Write-Host (Get-DistributionGroup $correspondantinformatique).OrganizationalUnit
}
else
{
	Write-Host "Groupe de Distribution $correspondantinformatique Inconnu !"
}

# Membres du groupe de distribution Correspondant Informatique

$chefdecentre = $dept + "chf"
$adjoint = $dept + "adj"
Write-Host $chefdecentre
Write-Host $adjoint

if (((Get-DistributionGroup -Identity $chefdecentre -ErrorAction 'SilentlyContinue').IsValid) -eq $true)
{

	if ((Get-DistributionGroupMember $correspondantinformatique | select -expand distinguishedname) -contains (Get-DistributionGroup $chefdecentre).distinguishedname)
	{
		Write-Host "membre $chefdecentre dans groupe $correspondantinformatique"
	}
	else
	{
		Write-Host "ATTENTION!!!! $chefdecentre n est pas membre du groupe $correspondantinformatique"
	}	
}

if (((Get-DistributionGroup -Identity $adjoint -ErrorAction 'SilentlyContinue').IsValid) -eq $true)
{

	if ((Get-DistributionGroupMember $correspondantinformatique | select -expand distinguishedname) -contains (Get-DistributionGroup $adjoint).distinguishedname)
	{
		Write-Host "membre $adjoint dans groupe $correspondantinformatique"
	}
	else
	{
		Write-Host "ATTENTION!!!! $adjoint n est pas membre du groupe $correspondantinformatique"

	}
}



# Correspondant Patrimoine

if (((Get-DistributionGroup -Identity $correspondantpatrimoine -ErrorAction 'SilentlyContinue').IsValid) -eq $true)
{
	Write-Host "Groupe de Distribution $correspondantpatrimoine OK !"
}
else
{
	Write-Host "Groupe de Distribution $correspondantpatrimoine Inconnu !"
}


#>
#Set-DistributionGroup -Identity "arlcinf" -alias "arlinf"
Remove-PSSession $session