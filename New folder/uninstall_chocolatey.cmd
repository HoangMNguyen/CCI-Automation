@echo off
:: Uninstall all Chocolatey packages
for /f "tokens=*" %%i in ('choco list --local-only --id-only') do choco uninstall -y %%i

:: Remove Chocolatey installation directory
rd /s /q %ALLUSERSPROFILE%\chocolatey

:: Remove Chocolatey from system PATH
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Environment]::SetEnvironmentVariable('PATH', (($env:PATH -split ';') -notmatch 'C:\ProgramData\chocolatey\bin') -join ';', [System.EnvironmentVariableTarget]::Machine)"

echo Chocolatey uninstalled successfully!
pause
