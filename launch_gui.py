#!/usr/bin/env python3
"""
盒标GUI启动脚本
"""

import sys
import os
from pathlib import Path

# 添加src目录到Python路径
current_dir = Path(__file__).parent
src_dir = current_dir / "src"
sys.path.insert(0, str(src_dir))

def main():
    """启动GUI应用程序"""
    try:
        print("🚀 启动盒标生成工具...")
        
        from gui.modern_ui import ModernExcelToPDFApp
        
        app = ModernExcelToPDFApp()
        print("✅ GUI已启动")
        app.run()
        
    except ImportError as e:
        print(f"❌ 导入错误: {e}")
        print("请确保所有依赖包已安装:")
        print("pip install pandas openpyxl reportlab pillow")
        sys.exit(1)
    except Exception as e:
        print(f"❌ 启动失败: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()