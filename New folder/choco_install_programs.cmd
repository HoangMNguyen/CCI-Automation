
@echo off
:: Refresh the environment variables
SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"

:: Install Docker CLI and Docker Engine
choco install -y wsl2
choco install -y docker-desktop
choco install -y docker-cli
choco install -y vcxsrv
choco install -y docker-compose

:: Add Docker to the system PATH
SETX PATH "%PATH%;C:\Program Files\Docker\Docker\resources\bin"

:: Verify Docker installation
docker --version
echo "Successfully installed programs."
choco list
pause