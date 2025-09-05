#!/usr/bin/env python3
"""
快速Windows兼容性测试
适合在Windows电脑上快速验证基本功能
"""

import sys
import os
import platform

def quick_test():
    """快速测试核心功能"""
    print(f"🖥️ 当前系统: {platform.system()} {platform.release()}")
    print(f"🐍 Python版本: {sys.version}")
    print("-" * 40)
    
    # 添加源码路径
    sys.path.insert(0, 'src')
    
    try:
        # 1. 测试核心导入
        print("1️⃣ 测试模块导入...")
        from pdf.generator import PDFGenerator
        from gui_app import DataToPDFApp
        print("   ✅ 核心模块导入成功")
        
        # 2. 测试字体
        print("2️⃣ 测试字体支持...")
        from utils.font_manager import font_manager
        font_name = font_manager.get_chinese_font_name()
        print(f"   ✅ 中文字体: {font_name}")
        
        # 3. 测试PDF生成器
        print("3️⃣ 测试PDF生成器...")
        generator = PDFGenerator()
        regular_template = generator.regular_template
        print("   ✅ PDF生成器创建成功")
        
        # 4. 测试文件路径
        print("4️⃣ 测试文件路径...")
        from pathlib import Path
        test_path = Path("test") / "folder" / "file.pdf"
        print(f"   ✅ 路径处理: {test_path}")
        
        print("\n🎉 快速测试完成！基本功能正常")
        return True
        
    except Exception as e:
        print(f"\n❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("🚀 Windows快速兼容性测试")
    print("=" * 40)
    
    success = quick_test()
    
    if success:
        print("\n✨ 建议下一步: 运行完整GUI程序")
        print("   命令: python src/gui_app.py")
    else:
        print("\n⚠️ 请检查Python环境和依赖包安装")
        print("   安装依赖: pip install -r requirements.txt")
    
    input("\n按Enter键退出...")