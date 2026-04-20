#!/bin/bash
# 文档管家 - Linux/麒麟 安装依赖脚本
# 使用方法: chmod +x install.sh && ./install.sh

echo "========================================"
echo "   文档管家 - 依赖安装脚本 (Linux/麒麟)"
echo "========================================"
echo ""

# 检查 Python3
if ! command -v python3 &> /dev/null; then
    echo "[错误] 未检测到 python3，请先安装 Python 3.8+"
    echo "  Ubuntu/Debian: sudo apt install python3 python3-pip python3-tk"
    echo "  麒麟系统: sudo apt install python3 python3-pip python3-tk"
    exit 1
fi

echo "[1/4] Python3 版本: $(python3 --version)"

# 检查 tkinter
python3 -c "import tkinter" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "[警告] tkinter 未安装，正在尝试安装..."
    sudo apt install -y python3-tk 2>/dev/null || {
        echo "[错误] tkinter 安装失败，请手动安装"
        exit 1
    }
fi
echo "[2/4] tkinter: 已就绪"

# 安装 Python 依赖
echo "[3/4] 正在安装 Python 依赖包..."
pip3 install -r requirements.txt --break-system-packages 2>/dev/null || \
pip3 install -r requirements.txt --user 2>/dev/null || \
pip install -r requirements.txt --break-system-packages 2>/dev/null

echo "[4/4] 安装完成！"
echo ""
echo "========================================"
echo "  启动方式: ./start.sh 或 python3 main.py"
echo "========================================"
