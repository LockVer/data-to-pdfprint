#!/bin/bash
echo "开始构建macOS/Linux版本..."
python3 build_config.py
if [ $? -eq 0 ]; then
    echo "构建成功！"
    echo "可执行文件位置: dist/PDFLabelGenerator/"
else
    echo "构建失败！"
fi
read -p "按任意键继续..."
