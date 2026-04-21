@echo off
chcp 65001 >nul
title 文档管家 - 安装依赖
echo ========================================
echo    文档管家 - 自动安装依赖
echo ========================================
echo.

echo [1/2] 正在检查并安装Python依赖包...
pip install python-docx openpyxl PyPDF2 pdfplumber python-pptx olefile chardet 2>nul
echo.

echo [2/2] 依赖安装完成，正在启动文档管家...
echo.
start "" pythonw main.py
