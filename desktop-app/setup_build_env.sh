#!/bin/bash

echo "🔧 设置构建环境..."

# 检查Python版本
python3 --version
if [ $? -ne 0 ]; then
    echo "❌ 未找到Python3，请先安装Python"
    exit 1
fi

# 升级pip
echo "📦 升级pip..."
python3 -m pip install --upgrade pip

# 安装依赖
echo "📚 安装项目依赖..."
python3 -m pip install -r requirements.txt

echo "✅ 构建环境设置完成！"
echo ""
echo "现在可以运行构建脚本："
echo "  macOS/Linux: ./build_unix.sh"
echo "  或手动: python3 build_config.py"