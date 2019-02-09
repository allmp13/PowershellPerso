@Echo off
echo %~dp0
echo %CD%
cd /d %~dp0
powershell -ExecutionPolicy ByPass -File ./Tailleprofils.ps1
copy ../NOK.txt ../Servers.txt
del ../NOK.txt
pause
popd