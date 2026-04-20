#!/bin/bash
# 文档管家 - Linux/麒麟 启动脚本（自动检测安装依赖）
# 使用方法: chmod +x start.sh && ./start.sh

cd "$(dirname "$0")"

# 检查 python3
if command -v python3 &> /dev/null; then
    PYTHON=python3
elif command -v python &> /dev/null; then
    PYTHON=python
else
    echo "错误: 未找到 Python，请先安装 Python 3.8+"
    echo "  麒麟系统: sudo apt install python3 python3-pip python3-tk"
    exit 1
fi

echo "Python: $($PYTHON --version 2>&1)"

# 检查 tkinter
$PYTHON -c "import tkinter" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "正在安装 tkinter..."
    sudo apt install -y python3-tk 2>/dev/null
fi

# 自动检测并安装缺失的 Python 依赖
echo "正在检查依赖..."
INSTALLED=0
FAILED=0

install_pkg() {
    local module=$1
    local pkg=$2
    $PYTHON -c "import $module" 2>/dev/null
    if [ $? -ne 0 ]; then
        echo "  安装 $pkg ..."
        pip3 install "$pkg" --break-system-packages 2>/dev/null || \
        pip3 install "$pkg" --user 2>/dev/null || \
        pip install "$pkg" --break-system-packages 2>/dev/null
        # 验证安装
        $PYTHON -c "import $module" 2>/dev/null
        if [ $? -eq 0 ]; then
            echo "    [OK] $module 安装成功"
            INSTALLED=$((INSTALLED+1))
        else
            echo "    [跳过] $module 安装失败（将使用内置备用方案）"
            FAILED=$((FAILED+1))
        fi
    else
        echo "  [OK] $module 已安装"
    fi
}

install_pkg "docx" "python-docx>=0.8.11"
install_pkg "openpyxl" "openpyxl>=3.1.0"
install_pkg "PyPDF2" "PyPDF2>=3.0.0"
install_pkg "pdfplumber" "pdfplumber>=0.10.0"
install_pkg "pptx" "python-pptx>=0.6.21"
install_pkg "olefile" "olefile>=0.46"
install_pkg "chardet" "chardet>=5.0.0"

echo ""
echo "依赖检查完成: 新安装 $INSTALLED 个, 失败 $FAILED 个"
if [ $FAILED -gt 0 ]; then
    echo "提示: 部分依赖安装失败，但程序仍可运行（使用内置备用提取方案）"
fi

# 启动应用
echo ""
echo "正在启动文档管家..."
$PYTHON main.py
