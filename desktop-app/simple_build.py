#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简化的构建脚本 - 使用PyInstaller快速打包
"""

import os
import sys
import subprocess
import platform
from pathlib import Path

def run_simple_build():
    """运行简化的构建流程"""
    
    current_dir = Path(__file__).parent
    root_dir = current_dir.parent
    
    print("🚀 开始简化构建...")
    print(f"📱 平台: {platform.system()}")
    
    # 基础PyInstaller命令
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=PDFLabelGenerator",
        "--onedir",
        "--windowed",
        "--noconfirm",
        "--clean",
        
        # 添加数据文件
        f"--add-data={root_dir}/src/fonts{os.pathsep}fonts",
        
        # 隐藏导入
        "--hidden-import=openpyxl",
        "--hidden-import=pandas", 
        "--hidden-import=reportlab",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        "--hidden-import=PIL",
        "--hidden-import=core",
        "--hidden-import=gui",
        "--hidden-import=gui.components",
        
        "main.py"
    ]
    
    try:
        print("🔨 执行构建命令...")
        print(f"命令: {' '.join(cmd)}")
        
        result = subprocess.run(cmd, cwd=current_dir, check=True)
        
        print("✅ 构建完成！")
        
        # 检查输出
        dist_dir = current_dir / "dist" / "PDFLabelGenerator"
        if dist_dir.exists():
            print(f"📦 输出目录: {dist_dir}")
            
            # 计算大小
            total_size = 0
            for root, dirs, files in os.walk(dist_dir):
                for file in files:
                    total_size += os.path.getsize(os.path.join(root, file))
            
            print(f"📊 总大小: {total_size / (1024*1024):.1f} MB")
            
            exe_name = "PDFLabelGenerator.exe" if platform.system() == "Windows" else "PDFLabelGenerator"
            exe_path = dist_dir / exe_name
            if exe_path.exists():
                print(f"🎯 可执行文件: {exe_path}")
                print(f"✨ 可以直接运行: {exe_path}")
        
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ 构建失败: {e}")
        return False
    except FileNotFoundError:
        print("❌ 未找到PyInstaller，请先安装:")
        print("   pip install pyinstaller")
        return False

if __name__ == "__main__":
    success = run_simple_build()
    if not success:
        input("按回车键退出...")
    sys.exit(0 if success else 1)