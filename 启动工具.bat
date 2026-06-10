@echo off
chcp 65001 >nul
title MinerU 文档批量处理工具
echo 正在启动 MinerU 文档批量处理工具...
python "%~dp0MinerU_Batch_Processor.py"
if %errorlevel% neq 0 (
    echo.
    echo 启动失败！请确保已安装 Python 和 mineru-open-sdk
    echo 安装命令: pip install mineru-open-sdk
    pause
)
