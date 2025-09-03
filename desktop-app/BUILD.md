# 桌面应用打包指南

这个文档说明如何将PDF标签生成器打包成Windows和macOS可执行文件。

## 系统要求

### 通用要求
- Python 3.8 或更高版本
- 安装所有依赖包：`pip install -r requirements.txt`

### Windows
- Windows 10 或更高版本
- 推荐使用 Windows PowerShell 或 Command Prompt

### macOS  
- macOS 10.14 或更高版本
- 已安装 Xcode Command Line Tools

## 一键构建（推荐）

### 第一步：设置环境

**Windows:**
```bash
# 双击运行或在命令行执行
setup_build_env.bat
```

**macOS/Linux:**
```bash
./setup_build_env.sh
```

### 第二步：开始构建

**简化构建（推荐）:**
```bash
python simple_build.py
```

**完整构建:**
- Windows: 双击 `build_windows.bat`  
- macOS/Linux: 执行 `./build_unix.sh`

## 快速构建

### 自动构建（推荐）

**Windows:**
```bash
# 双击运行或在命令行执行
build_windows.bat
```

**macOS/Linux:**
```bash
# 在终端中执行
./build_unix.sh
```

### 手动构建

1. **安装依赖**
```bash
pip install -r requirements.txt
```

2. **执行构建**
```bash
python build_config.py
```

## 构建输出

构建成功后，可执行文件将位于：
```
dist/PDFLabelGenerator/
├── PDFLabelGenerator.exe    # Windows可执行文件
├── PDFLabelGenerator        # macOS/Linux可执行文件  
├── fonts/                   # 字体文件
└── 其他依赖文件...
```

## 分发说明

### Windows分发
- 将整个 `dist/PDFLabelGenerator/` 文件夹打包为ZIP文件
- 用户解压后直接运行 `PDFLabelGenerator.exe`

### macOS分发  
- 将整个 `dist/PDFLabelGenerator/` 文件夹打包为ZIP文件
- 用户解压后运行 `PDFLabelGenerator` 文件
- 可能需要在"系统偏好设置 > 安全性与隐私"中允许运行

## 故障排除

### 常见问题

1. **构建失败：找不到模块**
   - 确保安装了所有依赖：`pip install -r requirements.txt`
   - 检查Python版本是否符合要求

2. **字体文件缺失**
   - 确保 `../src/fonts/` 目录存在且包含字体文件
   - 检查构建配置中的路径是否正确

3. **可执行文件运行失败**
   - 在终端中运行可执行文件查看错误信息
   - 检查是否缺少系统依赖

4. **文件过大**
   - 构建配置已排除不必要的模块
   - 可以考虑使用 `--onefile` 选项创建单文件版本

### 高级配置

如需自定义构建配置，可以编辑 `PDFLabelGenerator.spec` 文件：

- 修改包含的数据文件
- 添加或排除模块
- 自定义图标和版本信息
- 调整压缩选项

## 开发说明

### 目录结构
```
desktop-app/
├── main.py              # 应用入口
├── core/                # 核心数据模块
├── gui/                 # GUI组件
├── requirements.txt     # 依赖列表
├── build_config.py      # 构建配置脚本
├── build_windows.bat    # Windows构建脚本
├── build_unix.sh        # Unix构建脚本
└── BUILD.md            # 构建说明文档
```

### 添加新依赖
1. 在 `requirements.txt` 中添加包
2. 在 `build_config.py` 的 `hiddenimports` 中添加隐藏导入
3. 重新构建应用

### 自定义图标
1. 准备图标文件：
   - Windows: `icon.ico` (32x32, 48x48, 256x256)
   - macOS: `icon.icns` (多分辨率)
2. 将图标文件放在 `desktop-app/` 目录
3. 重新构建应用