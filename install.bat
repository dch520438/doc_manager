@echo off
chcp 65001 >nul
title 文档管家 - 启动（自动安装依赖）
echo ========================================
echo    文档管家 - 自动安装依赖并启动
echo ========================================
echo.

echo [1/2] 正在检查并安装Python依赖包...
pip install python-docx openpyxl PyPDF2 pdfplumber python-pptx olefile chardet --break-system-packages 2>nul
echo.

echo [2/2] 正在启动文档管家...
echo.
python main.py
if %errorlevel% neq 0 (
    echo.
    echo 启动失败，请检查是否已安装 Python 3.8+
    pause
)
