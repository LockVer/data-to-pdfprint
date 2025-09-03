#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF标签生成器 - 启动脚本
提供更好的错误处理和环境检查
"""

import sys
import os
import tkinter as tk
from tkinter import messagebox


def check_python_version():
    """检查Python版本"""
    if sys.version_info < (3, 8):
        print("错误: 需要Python 3.8或更高版本")
        print(f"当前版本: Python {sys.version}")
        return False
    return True


def check_dependencies():
    """检查必要的依赖包"""
    required_packages = [
        'tkinter',
        'reportlab', 
        'pandas',
        'openpyxl',
        'PIL'
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            if package == 'PIL':
                import PIL
            else:
                __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print("错误: 缺少必要的依赖包:")
        for package in missing_packages:
            print(f"  - {package}")
        print("\n请运行以下命令安装依赖:")
        print("pip install -r requirements.txt")
        return False
    
    return True


def check_project_structure():
    """检查项目结构"""
    required_dirs = [
        'gui',
        'core', 
        '../src'  # 需要访问原项目的src目录
    ]
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    for dir_name in required_dirs:
        dir_path = os.path.join(current_dir, dir_name)
        if not os.path.exists(dir_path):
            print(f"错误: 缺少必要目录 {dir_path}")
            return False
    
    return True


def main():
    """主函数"""
    print("PDF标签生成器 v1.0")
    print("正在启动...")
    
    # 环境检查
    if not check_python_version():
        input("按任意键退出...")
        return
    
    if not check_dependencies():
        input("按任意键退出...")
        return
    
    if not check_project_structure():
        input("按任意键退出...")
        return
    
    try:
        # 导入并启动应用
        from gui.main_window import PDFLabelGeneratorApp
        
        print("启动GUI界面...")
        app = PDFLabelGeneratorApp()
        app.run()
        
    except ImportError as e:
        error_msg = f"导入模块失败: {str(e)}"
        print(f"错误: {error_msg}")
        
        # 尝试显示GUI错误消息
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("启动错误", error_msg)
        except:
            pass
        
        input("按任意键退出...")
        
    except Exception as e:
        error_msg = f"应用启动失败: {str(e)}"
        print(f"错误: {error_msg}")
        
        # 尝试显示GUI错误消息
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("运行错误", error_msg)
        except:
            pass
        
        input("按任意键退出...")


if __name__ == "__main__":
    main()