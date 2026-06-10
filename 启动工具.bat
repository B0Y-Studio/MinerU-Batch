@echo off
chcp 65001 >nul
title MinerU Batch Processor
echo ====================================
echo   MinerU Document Batch Processor
echo ====================================
echo.

REM Try multiple Python paths
set PYTHON_CMD=
where python >nul 2>&1 && python -c "import mineru" >nul 2>&1 && set PYTHON_CMD=python
if "%PYTHON_CMD%"=="" (
    "D:\Softwares\python 3.11.7\python.exe" -c "import mineru" >nul 2>&1 && set PYTHON_CMD="D:\Softwares\python 3.11.7\python.exe"
)
if "%PYTHON_CMD%"=="" (
    py -3 -c "import mineru" >nul 2>&1 && set PYTHON_CMD=py -3
)

if "%PYTHON_CMD%"=="" (
    echo [!] Python not found or mineru-open-sdk not installed.
    echo     Install: pip install mineru-open-sdk
    pause
    exit /b 1
)

%PYTHON_CMD% "%~dp0MinerU_Batch_Processor.py"
if %errorlevel% neq 0 (
    echo.
    echo [!] Program exited with error code %errorlevel%
    pause
)
