Clear-Host
$password = Get-Content c:\scripts\MonFichier.pwd | ConvertTo-SecureString
$user = "ddsis-13\jmdavid"
$credential = New-Object System.Management.Automation.PSCredential($user,$password)
$session = New-PSSession -ConfigurationName Microsoft.Exchange –ConnectionUri http://srvcas1/powershell -Credential $credential
Import-PSSession $session -AllowClobber -DisableNameChecking
Import-Module ActiveDirectory

#Get-ADOrganizationalUnit -Filter 'Name -like "*"' | FT Name, DistinguishedName -A
#New-ADUser -SamAccountName "glenjohn" -GivenName "Glen" -Surname "John" -DisplayName "Glen John" -Path 'CN=Users,DC=fabrikam,DC=local'


#get-distributiongroup | select name, Hiddenfromaddresslistsenabled | WHERE {$_.HiddenFromAddressListsEnabled -eq $true }
#Get-DistributionGroup -FILTER {EmailAddresses -like "*GSI*"} | select name,primarysmtpaddress,RequireSenderAuthenticationEnabled,HiddenFromAddressListsEnabled,EmailAddresses
#Get-DistributionGroup -FILTER {name -like "correspondant informatique*"} | select name,primarysmtpaddress |  Export-Csv C:\distribution_group.csv


#(Get-DistributionGroup -Identity  "chefprogrammationsoutientechnique@sdis13.fr").RequireSenderAuthenticationEnabled
 #Write-Output $user | Out-File 'c:\Servers.txt' -Append

#Get-DistributionGroupMember -Identity aixinf@sdis13.fr
#New-DistributionGroup -OrganizationalUnit  "sdis13.fr/CS/Avec Serveur/AIX/Groupes/Distribution" -Alias "totoaix4" -Name "Toto Aix 4" 
#Set-DistributionGroup -Identity totoaix4 -RequireSenderAuthenticationEnabled $False
<#
$groupname="totoaix3"
if (((Get-DistributionGroup -Identity $groupname -ErrorAction 'SilentlyContinue').IsValid) -eq $true)
{
	#Get-DistributionGroup -Identity $groupname | Format-List
	Get-DistributionGroup -Identity $groupname | select displayname, Name, Alias, PrimarySmtpAddress
}
else
{
	Write-Host "Groupe de Distribution Inconnu !"
}

Get-DistributionGroupMember "Sceops" | Format-Table -property PrimarySmtpAddress
(Get-DistributionGroupMember "Sceops") | -contains "lleandri@sdis13.fr"

 $user1="aixinf@sdis13.fr"
 Get-DistributionGroup $user1
 $user_dn = (Get-DistributionGroup $user1).distinguishedname
 Write-Host $user_dn
 Write-Host (Get-DistributionGroup $user1).OrganizationalUnit
(Get-DistributionGroupMember "Sceops" | select -expand distinguishedname) -contains $user_dn
(Get-Recipient -Identity $user1).RecipientType
(Get-Recipient -Identity "jmdavid@sdis13.fr").RecipientType

Get-DistributionGroupMember -Identity aixinf@sdis13.fr
New-DistributionGroup -OrganizationalUnit  "sdis13.fr/CS/Avec Serveur/AIX/Groupes/Distribution" -Alias "totoaix4" -Name "Toto Aix 4" 
Set-DistributionGroup -Identity totoaix4 -RequireSenderAuthenticationEnabled $False

$groupname="totoaix3"
if (((Get-DistributionGroup -Identity $groupname -ErrorAction 'SilentlyContinue').IsValid) -eq $true)
{
	#Get-DistributionGroup -Identity $groupname | Format-List
	Get-DistributionGroup -Identity $groupname | select displayname, Name, Alias, PrimarySmtpAddress
}
else
{
	Write-Host "Groupe de Distribution Inconnu !"
}

Get-DistributionGroupMember "Sceops" | Format-Table -property PrimarySmtpAddress
(Get-DistributionGroupMember "Sceops") | -contains "lleandri@sdis13.fr"

 $user1="aixinf@sdis13.fr"
 $user_dn = (Get-DistributionGroup $user1).distinguishedname
 Write-Host $user_dn
(Get-DistributionGroupMember "Sceops" | select -expand distinguishedname) -contains $user_dn
(Get-Recipient -Identity $user1).RecipientType
(Get-Recipient -Identity "jmdavid@sdis13.fr").RecipientType
 #>
 

Remove-PSSession $session
