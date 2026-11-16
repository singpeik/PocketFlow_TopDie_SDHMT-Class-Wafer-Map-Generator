@echo off
title Wafer Test Result Analyzer - SDHMT and Class
setlocal

cd /d "%~dp0"

set "VENV_PATH=C:\venvs\pocketflow-IOE-SDHMT-Class-wafermap-generator"
set "VENV_ACTIVATE=%VENV_PATH%\Scripts\activate.bat"

if not exist "%VENV_ACTIVATE%" (
    echo [ERROR] Virtual environment activation script not found at:
    echo %VENV_ACTIVATE%
    echo.
    echo Please make sure the venv exists in C:\venvs.
    pause
    exit /b 1
)

call "%VENV_ACTIVATE%"
if errorlevel 1 (
    echo [ERROR] Failed to activate the virtual environment.
    pause
    exit /b 1
)

echo Starting Wafer Test Result Analyzer - SDHMT and Class...
echo Working directory: %cd%
echo Using Python: %VIRTUAL_ENV%\Scripts\python.exe
echo.

python main.py %*

if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] Script finished with error code %ERRORLEVEL%.
) else (
    echo.
    echo [SUCCESS] Script completed successfully.
)

pause
