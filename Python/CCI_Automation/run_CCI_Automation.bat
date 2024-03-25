@echo off

REM Activate the virtual environment
call "%~dp0\env\Scripts\activate"

REM Check if running in a virtual environment
if defined VIRTUAL_ENV (
    echo Running in virtual environment.
) else (
    echo Not running in a virtual environment. Please activate the virtual environment first.
    exit /b
)

REM Run the Python script
call "%~dp0CCI_Automation.py"

REM Deactivate virtual environment
deactivate
