@echo off
SETLOCAL

:: Check if Chocolatey is installed
WHERE choco >nul 2>&1
IF ERRORLEVEL 1 (
    @echo Chocolatey not found. Installing now...
    @powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"
)

:: Check if Python is installed
WHERE python >nul 2>&1
IF ERRORLEVEL 1 (
    @echo Python not found. Installing now...
    choco install python -y
) ELSE (
    @echo Python found. Upgrading now...
    choco upgrade python -y
)

PAUSE
@echo Press any key to exit

ENDLOCAL
