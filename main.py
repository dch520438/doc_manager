"""
文档管家 - 启动入口
双击运行此文件即可启动系统
"""

import os
import sys

# 确保可以导入同目录下的模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui import DocManagerApp

if __name__ == "__main__":
    app = DocManagerApp()
    app.run()
