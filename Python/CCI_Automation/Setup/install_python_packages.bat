@echo off

echo Installing Python...


echo Checking Python installation...
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo Python is already installed.
) else (
    echo Downloading Python version 3.11.2...
    curl -o python-3.11.2-amd64.exe https://www.python.org/ftp/python/3.11.2/python-3.11.2-amd64.exe
    echo Installing Python...
    start /wait python-3.11.2-amd64.exe /quiet InstallAllUsers=1 PrependPath=1
)

echo Checking pip installation...
pip --version >nul 2>&1
if %errorlevel% equ 0 (
    echo pip is already installed.
) else (
    echo Installing pip...
    curl -o get-pip.py https://bootstrap.pypa.io/get-pip.py
    echo Installing pip...
    python get-pip.py
)

echo Installing dependencies listed in requirements.txt...

for /F "tokens=*" %%G in (packages.txt) do (
    pip show %%G >nul 2>&1
    echo Installing %%G...
    pip install %%G
    )
)


echo Cleaning up...
del get-pip.py
del python-3.11.2-amd64.exe

echo Installation complete. No error? Hoang is awesome!

pause