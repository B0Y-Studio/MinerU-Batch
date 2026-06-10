@echo off
chcp 65001 >nul
title MinerU Batch Processor
echo ====================================
echo   MinerU Document Batch Processor
echo ====================================
echo.
python "%~dp0MinerU_Batch_Processor.py"
if %errorlevel% neq 0 (
    echo.
    echo [!] Failed to start. Make sure Python and mineru-open-sdk are installed.
    echo     Install: pip install mineru-open-sdk
    pause
)
