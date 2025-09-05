#!/usr/bin/env python3
"""
Windows兼容性测试脚本
用于验证项目在Windows系统上的功能
"""

import sys
import os
import platform
import traceback
from pathlib import Path

def test_basic_imports():
    """测试基本模块导入"""
    print("🔍 测试基本模块导入...")
    
    try:
        sys.path.insert(0, 'src')
        
        # 测试核心模块
        from pdf.generator import PDFGenerator
        print("  ✅ PDFGenerator导入成功")
        
        from gui_app import DataToPDFApp
        print("  ✅ DataToPDFApp导入成功")
        
        from utils.excel_data_extractor import ExcelDataExtractor
        print("  ✅ ExcelDataExtractor导入成功")
        
        from utils.font_manager import font_manager
        print("  ✅ font_manager导入成功")
        
        return True
    except Exception as e:
        print(f"  ❌ 导入失败: {e}")
        traceback.print_exc()
        return False

def test_font_detection():
    """测试Windows字体检测"""
    print("\n🔍 测试Windows字体检测...")
    
    try:
        sys.path.insert(0, 'src')
        from utils.font_manager import font_manager
        
        # 检查平台
        current_platform = platform.system()
        print(f"  📱 当前平台: {current_platform}")
        
        # 测试字体注册
        font_name = font_manager.get_chinese_font_name()
        print(f"  ✅ 中文字体: {font_name}")
        
        return True
    except Exception as e:
        print(f"  ❌ 字体检测失败: {e}")
        traceback.print_exc()
        return False

def test_pdf_generator_creation():
    """测试PDF生成器创建"""
    print("\n🔍 测试PDF生成器创建...")
    
    try:
        sys.path.insert(0, 'src')
        from pdf.generator import PDFGenerator
        
        # 创建生成器实例
        generator = PDFGenerator()
        print("  ✅ PDFGenerator实例创建成功")
        
        # 测试延迟加载模板
        regular = generator.regular_template
        print("  ✅ 常规模板延迟加载成功")
        
        split_box = generator.split_box_template  
        print("  ✅ 分盒模板延迟加载成功")
        
        nested_box = generator.nested_box_template
        print("  ✅ 套盒模板延迟加载成功")
        
        return True
    except Exception as e:
        print(f"  ❌ PDF生成器测试失败: {e}")
        traceback.print_exc()
        return False

def test_file_paths():
    """测试文件路径处理"""
    print("\n🔍 测试文件路径处理...")
    
    try:
        # 测试路径创建
        test_path = Path("output") / "test_folder" / "test_file.pdf"
        print(f"  📁 路径测试: {test_path}")
        
        # 测试文件夹名称清理
        test_theme = "TEST/FILE*NAME<WITH>SPECIAL|CHARS"
        clean_theme = test_theme.replace('\n', ' ').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('*', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').replace('!', '_')
        print(f"  🧹 清理前: {test_theme}")
        print(f"  🧹 清理后: {clean_theme}")
        
        return True
    except Exception as e:
        print(f"  ❌ 路径测试失败: {e}")
        traceback.print_exc()
        return False

def test_excel_processing():
    """测试Excel处理（如果有测试文件）"""
    print("\n🔍 测试Excel处理...")
    
    # 检查是否有测试Excel文件
    test_files = ["test.xlsx", "data/test.xlsx", "samples/test.xlsx"]
    excel_file = None
    
    for file_path in test_files:
        if Path(file_path).exists():
            excel_file = file_path
            break
    
    if not excel_file:
        print("  ⚠️ 未找到测试Excel文件，跳过Excel处理测试")
        return True
    
    try:
        sys.path.insert(0, 'src')
        from utils.excel_data_extractor import ExcelDataExtractor
        
        extractor = ExcelDataExtractor(excel_file)
        data = extractor.extract_common_data()
        print(f"  ✅ Excel数据提取成功: {list(data.keys())}")
        
        return True
    except Exception as e:
        print(f"  ❌ Excel处理失败: {e}")
        traceback.print_exc()
        return False

def main():
    """主测试函数"""
    print("🚀 Windows兼容性测试开始")
    print("=" * 50)
    
    tests = [
        ("基本模块导入", test_basic_imports),
        ("字体检测", test_font_detection),
        ("PDF生成器", test_pdf_generator_creation),
        ("文件路径", test_file_paths),
        ("Excel处理", test_excel_processing),
    ]
    
    results = []
    
    for test_name, test_func in tests:
        try:
            success = test_func()
            results.append((test_name, success))
        except Exception as e:
            print(f"💥 测试 '{test_name}' 出现异常: {e}")
            results.append((test_name, False))
    
    # 汇总结果
    print("\n" + "=" * 50)
    print("🎯 测试结果汇总:")
    
    passed = 0
    total = len(results)
    
    for test_name, success in results:
        status = "✅ 通过" if success else "❌ 失败"
        print(f"  {test_name}: {status}")
        if success:
            passed += 1
    
    print(f"\n📊 总体结果: {passed}/{total} 项测试通过")
    
    if passed == total:
        print("🎉 所有测试通过！项目在当前环境下运行正常")
        return 0
    else:
        print("⚠️ 部分测试失败，需要进一步检查")
        return 1

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)