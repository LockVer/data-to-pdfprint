#!/usr/bin/env python3
"""
现代化GUI启动脚本

启动美化版本的Excel转PDF工具
"""

import sys
from pathlib import Path

# 添加src目录到Python路径
current_dir = Path(__file__).parent
src_dir = current_dir / 'src'
sys.path.insert(0, str(src_dir))

if __name__ == "__main__":
    try:
        from gui.modern_ui import main
        print("🚀 启动现代化Excel转PDF工具...")
        main()
    except ImportError as e:
        print(f"导入错误: {e}")
        print("请确保所有依赖包已安装:")
        print("pip install pandas openpyxl reportlab pillow")
    except Exception as e:
        print(f"启动错误: {e}")