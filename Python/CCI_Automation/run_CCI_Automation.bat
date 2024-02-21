@echo off

REM Activate the virtual environment
call "%~dp0\env\Scripts\activate"


REM Run the Python script
call "%~dp0CCI_Automation.py"

REM Deactivate virtual environment
deactivate
