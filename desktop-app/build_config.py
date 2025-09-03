#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
桌面应用打包配置和构建脚本
支持Windows和macOS平台
"""

import os
import sys
import platform
import subprocess
from pathlib import Path

# 项目基本信息
APP_NAME = "PDF标签生成器"
APP_NAME_EN = "PDFLabelGenerator"
VERSION = "1.0.0"
AUTHOR = "Your Company"
DESCRIPTION = "Excel数据到PDF标签生成工具"

# 获取当前目录
CURRENT_DIR = Path(__file__).parent
ROOT_DIR = CURRENT_DIR.parent

def get_pyinstaller_args():
    """获取PyInstaller打包参数"""
    
    # 基本参数
    args = [
        "pyinstaller",
        "--name", APP_NAME_EN,
        "--onedir",  # 创建单文件夹，包含所有依赖
        "--windowed",  # Windows下不显示控制台窗口
        "--noconfirm",  # 覆盖已存在的构建
        "--clean",  # 清理临时文件
    ]
    
    # 添加数据文件和隐藏导入
    args.extend([
        # 包含字体文件
        f"--add-data={ROOT_DIR}/src/fonts{os.pathsep}fonts",
        
        # 隐藏导入的模块
        "--hidden-import=openpyxl",
        "--hidden-import=pandas",
        "--hidden-import=reportlab",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        "--hidden-import=PIL",
        
        # 排除不需要的模块以减小文件大小
        "--exclude-module=matplotlib",
        "--exclude-module=numpy.testing",
        "--exclude-module=pandas.tests",
    ])
    
    # 根据平台添加特定参数
    if platform.system() == "Windows":
        args.extend([
            "--icon=icon.ico",  # Windows图标
        ])
    elif platform.system() == "Darwin":  # macOS
        args.extend([
            "--icon=icon.icns",  # macOS图标
        ])
    
    # 指定入口文件
    args.append(str(CURRENT_DIR / "main.py"))
    
    return args

def create_spec_file():
    """创建PyInstaller spec文件以便自定义配置"""
    
    spec_content = f'''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['{CURRENT_DIR}'],
    binaries=[],
    datas=[
        ('{ROOT_DIR}/src/fonts', 'fonts'),
    ],
    hiddenimports=[
        'openpyxl',
        'pandas',
        'reportlab',
        'tkinter',
        'tkinter.ttk',
        'PIL',
        'core',
        'gui',
        'gui.components',
    ],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy.testing',
        'pandas.tests',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='{APP_NAME_EN}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='{APP_NAME_EN}',
)
'''

    spec_file = CURRENT_DIR / f"{APP_NAME_EN}.spec"
    with open(spec_file, 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    return spec_file

def build_app():
    """构建应用程序"""
    
    print(f"🚀 开始构建 {APP_NAME} v{VERSION}")
    print(f"📱 目标平台: {platform.system()} {platform.machine()}")
    print(f"📂 工作目录: {CURRENT_DIR}")
    
    # 检查依赖
    try:
        import pyinstaller
        print(f"✅ PyInstaller 版本: {pyinstaller.__version__}")
    except ImportError:
        print("❌ 未找到PyInstaller，请先安装：pip install pyinstaller")
        return False
    
    # 创建spec文件
    print("📝 创建PyInstaller配置文件...")
    spec_file = create_spec_file()
    print(f"✅ 配置文件已创建: {spec_file}")
    
    # 执行构建
    try:
        print("🔨 开始构建...")
        cmd = ["pyinstaller", str(spec_file)]
        result = subprocess.run(cmd, cwd=CURRENT_DIR, check=True, capture_output=True, text=True)
        
        print("✅ 构建成功!")
        
        # 输出构建结果信息
        dist_dir = CURRENT_DIR / "dist" / APP_NAME_EN
        if dist_dir.exists():
            print(f"📦 构建输出目录: {dist_dir}")
            
            # 计算文件大小
            total_size = sum(f.stat().st_size for f in dist_dir.rglob('*') if f.is_file())
            print(f"📊 总大小: {total_size / (1024*1024):.1f} MB")
            
            # 列出主要文件
            exe_name = APP_NAME_EN + ('.exe' if platform.system() == 'Windows' else '')
            exe_path = dist_dir / exe_name
            if exe_path.exists():
                print(f"🎯 可执行文件: {exe_path}")
            
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ 构建失败: {e}")
        if e.stdout:
            print("标准输出:", e.stdout)
        if e.stderr:
            print("错误输出:", e.stderr)
        return False

def create_build_scripts():
    """创建跨平台构建脚本"""
    
    # Windows批处理脚本
    windows_script = '''@echo off
echo 开始构建Windows版本...
python build_config.py
if %ERRORLEVEL% EQU 0 (
    echo 构建成功！
    echo 可执行文件位置: dist\\PDFLabelGenerator\\
    pause
) else (
    echo 构建失败！
    pause
)
'''
    
    with open(CURRENT_DIR / "build_windows.bat", 'w', encoding='utf-8') as f:
        f.write(windows_script)
    
    # macOS/Linux shell脚本
    unix_script = '''#!/bin/bash
echo "开始构建macOS/Linux版本..."
python3 build_config.py
if [ $? -eq 0 ]; then
    echo "构建成功！"
    echo "可执行文件位置: dist/PDFLabelGenerator/"
else
    echo "构建失败！"
fi
read -p "按任意键继续..."
'''
    
    unix_script_path = CURRENT_DIR / "build_unix.sh"
    with open(unix_script_path, 'w', encoding='utf-8') as f:
        f.write(unix_script)
    
    # 给shell脚本添加执行权限
    os.chmod(unix_script_path, 0o755)
    
    print("✅ 构建脚本已创建:")
    print(f"   Windows: {CURRENT_DIR / 'build_windows.bat'}")
    print(f"   macOS/Linux: {unix_script_path}")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--create-scripts":
        create_build_scripts()
    else:
        success = build_app()
        sys.exit(0 if success else 1)