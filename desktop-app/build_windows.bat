@echo off
echo 开始构建Windows版本...
python build_config.py
if %ERRORLEVEL% EQU 0 (
    echo 构建成功！
    echo 可执行文件位置: dist\PDFLabelGenerator\
    pause
) else (
    echo 构建失败！
    pause
)
