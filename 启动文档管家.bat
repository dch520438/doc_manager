@echo off
chcp 65001 >nul
title 文档管家
cd /d "%~dp0"
python main.py
if %errorlevel% neq 0 (
    echo.
    echo 启动失败！请先运行 install.bat 安装依赖
    echo.
    pause
)
