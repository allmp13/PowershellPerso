#Exemples de gestion d'erreurs
#Version du 18 avril 2013

Clear-Host

$Error.Clear()
#Utilisation d'une variable qui agit au niveau du reste du script
#$ErrorActionPreference = "silentlycontinue" #Ne pas afficher les erreurs et continuer en cas de problème
#$ErrorActionPreference = "Stop" #Arrêter en cas d'erreur
#$ErrorActionPreference = "Inquire" #Demander quoi faire en cas de problème
#http://blogs.technet.com/b/heyscriptingguy/archive/2010/03/08/hey-scripting-guy-march-8-2010.aspx

#Ajouter -ErrorAction SilentlyContinue au niveau d'un appel n'affecte le comportement en cas d'erreur que lors de cet appel
#Ajouter -ErrorVariable ErrObj permet de récupérer un objet nommé ErrObj qui contient les infos de l'erreur générée
#Get-Content "C:\non-existent folder"
#Get-Content "C:\non-existent folder" -ErrorAction SilentlyContinue
Get-Content "C:\temp" -ErrorVariable ErrObj -ErrorAction SilentlyContinue

if ($error.Count -ieq 0){
	Write-Host "Tout est Ok"
}
else{
	#Affichage de l'erreur stockée dans $ErrObj
	Write-Host $ErrObj[0]
}

#A l'inverse on aurait pu utiliser : Si il y a une erreur :
#if($Error.Count -ne 0) {
#}