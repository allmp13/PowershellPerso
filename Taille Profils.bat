@Echo off
cd c:\scripts
powershell -ExecutionPolicy ByPass -File ./Tailleprofils.ps1
copy NOK.txt Servers.txt
del NOK.txt
pause