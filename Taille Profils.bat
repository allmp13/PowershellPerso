@Echo off
echo %~dp0
echo %CD%
pushd %~dp0
powershell -ExecutionPolicy ByPass -File ./Tailleprofils.ps1
copy ..\NOK.txt ..\Servers.txt /y
del ..\NOK.txt
pause
popd