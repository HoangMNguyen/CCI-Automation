@echo off
setlocal

:: Verify Docker installation by running hello-world container
echo Running hello-world container to verify Docker installation...
docker run hello-world
if %errorlevel% neq 0 (
    echo Failed to run hello-world container. Exiting...
    pause
    exit /b 1
)

echo Docker is installed and working correctly!
endlocal
pause
