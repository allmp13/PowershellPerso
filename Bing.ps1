﻿$nbdouble=0
for ($cpt=1;$cpt -le 66; $cpt++)
{
if ($nbdouble -gt '1000') { BREAK} 
#write-host $cpt
$page='http://www.binghomepagewallpapers.x10host.com/BingPost-'+$cpt.tostring("00")+'.html'
#write-host $page
#Read-Host "Press ENTER"
$HTML = ((Invoke-WebRequest -Uri $page).links  | Where-Object {$_.href -like “*1920x1080*”} )
$nomimg="http://www.binghomepagewallpapers.x10host.com"
$FolderToCreate="C:\Users\jmdavid\Pictures\wallpaper\"
ForEach ($toto in $HTML)
{
if ($nbdouble -gt '1000') { BREAK} 
#write-host $toto.href
$nom= $toto.title
$toto=$toto.href
#write-host $toto
$url=$nomimg + $toto.substring(1,$toto.Length-1)
#write-host $url
$nomfic=$FolderToCreate + $nom + ".jpg"
#write-host $nomfic

if (!(Test-Path -path $nomfic))
 
 {
  write-host  "Téléchargement de "  $nom
 (New-Object System.Net.WebClient).DownloadFile($url, $nomfic)
 $nbdouble=0
 }

 else {
 
         write-host $nom "Existe deja"
         $nbdouble++
         #write-host $nbdouble
       }
}
}



