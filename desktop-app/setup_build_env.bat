@echo off
echo 🔧 设置构建环境...

:: 检查Python
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ❌ 未找到Python，请先安装Python
    pause
    exit /b 1
)

:: 升级pip
echo 📦 升级pip...
python -m pip install --upgrade pip

:: 安装依赖
echo 📚 安装项目依赖...
python -m pip install -r requirements.txt

echo ✅ 构建环境设置完成！
echo.
echo 现在可以运行构建脚本：
echo   Windows: build_windows.bat
echo   或手动: python build_config.py
pause