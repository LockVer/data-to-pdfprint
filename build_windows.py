"""
专门用于Windows系统的GUI构建脚本
确保生成的可执行文件能在Windows上正常运行
"""

import subprocess
import sys
import os
import platform
from pathlib import Path

def build_windows_gui():
    """构建Windows GUI应用"""
    
    if platform.system() != "Windows":
        print("⚠️  警告: 当前不是Windows系统，生成的可执行文件可能无法在Windows上运行")
        print("推荐在Windows系统上运行此脚本")
        
    print("正在构建Windows GUI应用...")
    
    # 确保在项目根目录
    project_root = Path(__file__).parent
    os.chdir(project_root)
    
    # 检查必要文件
    if not os.path.exists("src/gui_app.py"):
        print("❌ 错误: 找不到 src/gui_app.py")
        return
    
    # 使用专门的Windows配置文件
    cmd = [
        "pyinstaller",
        "DataToPDF_GUI_Windows.spec",
        "--clean",  # 清理缓存
        "--noconfirm",  # 不询问覆盖
    ]
    
    try:
        # 清理旧文件
        if os.path.exists("dist"):
            import shutil
            shutil.rmtree("dist")
        if os.path.exists("build"):
            import shutil
            shutil.rmtree("build")
        
        print("🔄 正在运行PyInstaller...")
        result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
        
        if result.returncode == 0:
            print("✅ Windows GUI应用构建成功!")
            print("📁 生成文件: dist/DataToPDF_GUI.exe")
            print("\n📋 Windows使用说明:")
            print("  - 双击 DataToPDF_GUI.exe 启动可视化界面")
            print("  - 支持Windows 7/8/10/11")
            print("  - 无需安装Python环境")
            print("  - 可以发送给其他Windows用户使用")
            print("\n🎯 操作流程:")
            print("1. 双击运行 DataToPDF_GUI.exe")
            print("2. 点击'选择Excel文件'按钮选择xlsx文件")
            print("3. 点击'生成PDF'按钮，选择保存位置")
            print("4. 自动提取总张数并生成标签PDF")
            print("\n💡 分发说明:")
            print("  - 可以直接复制 DataToPDF_GUI.exe 给其他用户")
            print("  - 建议打包成ZIP文件分发")
            print("  - 文件大小约 35-50MB")
            
        else:
            print("❌ 构建失败:")
            print("标准输出:", result.stdout)
            print("错误输出:", result.stderr)
            
            # 提供常见错误的解决建议
            if "No module named" in result.stderr:
                print("\n💡 解决建议:")
                print("请先安装依赖: pip install -r requirements.txt")
            elif "PyInstaller" not in result.stderr:
                print("\n💡 解决建议:")
                print("请先安装PyInstaller: pip install pyinstaller")
                
    except FileNotFoundError:
        print("❌ PyInstaller未找到，请先安装:")
        print("pip install pyinstaller")
    except Exception as e:
        print(f"❌ 构建过程出错: {e}")

def create_distribution_package():
    """创建Windows分发包"""
    if os.path.exists("dist/DataToPDF_GUI.exe"):
        print("\n📦 创建分发包...")
        
        # 创建分发目录
        dist_dir = Path("windows_distribution")
        if dist_dir.exists():
            import shutil
            shutil.rmtree(dist_dir)
        dist_dir.mkdir()
        
        # 复制可执行文件
        import shutil
        shutil.copy2("dist/DataToPDF_GUI.exe", dist_dir / "DataToPDF_GUI.exe")
        
        # 创建使用说明
        readme_content = """# Data to PDF Print - Windows版本

## 使用方法
1. 双击运行 DataToPDF_GUI.exe
2. 选择Excel文件（.xlsx格式）
3. 点击生成PDF按钮
4. 选择保存位置

## 系统要求
- Windows 7/8/10/11
- 无需安装Python或其他依赖

## 注意事项
- 首次运行可能被Windows Defender检测，选择"允许"即可
- 建议将程序放在非系统盘（如D盘）运行
- 支持的Excel格式：.xlsx, .xls

## 问题反馈
如有问题请联系开发者
"""
        
        with open(dist_dir / "使用说明.txt", "w", encoding="utf-8") as f:
            f.write(readme_content)
        
        print(f"✅ 分发包已创建: {dist_dir.absolute()}")
        print("📋 包含文件:")
        for file in dist_dir.iterdir():
            print(f"  - {file.name}")

if __name__ == "__main__":
    build_windows_gui()
    create_distribution_package()