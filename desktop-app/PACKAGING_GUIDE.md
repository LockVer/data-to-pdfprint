# PDF标签生成器 - 打包部署指南

## 🎯 概述

这是一个基于Python + tkinter的桌面应用程序，可以生成Windows和macOS的可执行文件。

## 📁 文件清单

### 核心文件
- `main.py` - 应用程序入口点
- `requirements.txt` - Python依赖包列表
- `core/` - 核心业务逻辑
- `gui/` - 用户界面组件

### 构建文件
- `simple_build.py` - 简化构建脚本（推荐）
- `build_config.py` - 高级构建配置脚本
- `build_windows.bat` - Windows一键构建
- `build_unix.sh` - macOS/Linux一键构建
- `setup_build_env.bat/.sh` - 环境设置脚本

### 说明文档
- `BUILD.md` - 详细构建说明
- `PACKAGING_GUIDE.md` - 本文档

## 🚀 快速开始

### 方式一：一键构建（最简单）

1. **设置环境**
   ```bash
   # Windows
   setup_build_env.bat
   
   # macOS/Linux
   ./setup_build_env.sh
   ```

2. **构建应用**
   ```bash
   python simple_build.py
   ```

### 方式二：使用构建脚本

```bash
# Windows
build_windows.bat

# macOS/Linux  
./build_unix.sh
```

## 📦 构建输出

构建成功后，文件结构如下：
```
dist/PDFLabelGenerator/
├── PDFLabelGenerator(.exe)  # 主程序
├── fonts/                   # 字体文件
├── _internal/              # Python运行时和依赖
└── 其他库文件...
```

## 🎁 分发说明

### Windows分发
1. 将整个 `dist/PDFLabelGenerator/` 文件夹打包为ZIP
2. 用户解压后直接运行 `PDFLabelGenerator.exe`
3. 无需安装Python环境

### macOS分发
1. 将整个 `dist/PDFLabelGenerator/` 文件夹打包为ZIP
2. 用户解压后运行 `PDFLabelGenerator` 文件
3. 首次运行可能需要在"系统偏好设置 > 安全性与隐私"中允许

## 🔧 技术细节

### 打包工具
- **PyInstaller** - Python应用打包工具
- **--onedir** - 创建包含所有依赖的文件夹
- **--windowed** - 不显示控制台窗口

### 包含的依赖
- tkinter - GUI框架
- reportlab - PDF生成
- pandas - Excel处理  
- openpyxl - Excel读写
- PIL/Pillow - 图像处理

### 文件大小
- Windows: 约 80-120MB
- macOS: 约 90-130MB

## ⚠️ 注意事项

1. **字体文件**：确保 `../src/fonts/` 目录存在
2. **权限问题**：macOS可能需要允许运行未签名应用
3. **防病毒软件**：某些杀毒软件可能误报，需要添加白名单
4. **系统兼容性**：
   - Windows 10 及以上
   - macOS 10.14 及以上

## 🐛 故障排除

### 常见问题

1. **PyInstaller未安装**
   ```bash
   pip install pyinstaller
   ```

2. **构建失败**
   - 检查Python版本 (需要3.8+)
   - 确保所有依赖已安装
   - 检查路径中是否有中文字符

3. **运行失败**
   - 在终端中运行查看错误信息
   - 检查是否缺少系统库

4. **文件过大**
   - 可以尝试使用 `--onefile` 创建单文件版本
   - 排除不必要的模块

### 调试技巧

1. **在控制台中运行**
   ```bash
   # Windows
   PDFLabelGenerator.exe
   
   # macOS  
   ./PDFLabelGenerator
   ```

2. **查看构建日志**
   - PyInstaller会生成详细的构建日志
   - 检查 `build/` 和 `dist/` 目录

## 📞 技术支持

如遇到构建或运行问题，请检查：
1. Python版本兼容性
2. 依赖包完整性
3. 系统权限设置
4. 防火墙/杀毒软件配置

---

💡 **提示**: 建议在干净的虚拟环境中进行构建，以避免依赖冲突。