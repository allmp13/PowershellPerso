@Echo off
cd C:\scripts\Powershell
powershell -ExecutionPolicy ByPass -File ./Tailleprofils.ps1
copy ../NOK.txt ../Servers.txt
del ../NOK.txt
pause