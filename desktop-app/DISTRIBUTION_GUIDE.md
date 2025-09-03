# PDF标签生成器 - 分发指南

## 📦 构建结果

### macOS 版本 ✅ 已完成
- **文件位置**: `/Users/qyx/Desktop/task/data-to-pdfprint/desktop-app/dist/PDFLabelGenerator-macOS.zip`
- **文件大小**: 62MB (压缩后)，解压后约 126MB
- **运行方式**: 解压后运行 `PDFLabelGenerator` 文件
- **系统要求**: macOS 10.14 及以上

### Windows 版本 🔄 待构建
- **构建环境**: 需要在 Windows 系统上执行
- **构建步骤**: 见下方 Windows 构建说明

## 🚀 快速分发

### macOS 用户
1. 下载 `PDFLabelGenerator-macOS.zip`
2. 解压到任意目录
3. 双击 `PDFLabelGenerator` 运行
4. 首次运行可能需要在"系统偏好设置 > 安全性与隐私"中允许

### Windows 用户
按照下方 Windows 构建说明在 Windows 环境中构建可执行文件。

## 🔨 Windows 构建说明

### 环境准备
在 Windows 机器上执行以下步骤：

1. **安装 Python 3.8+**
   ```cmd
   # 从 python.org 下载并安装 Python
   # 确保勾选 "Add Python to PATH"
   ```

2. **克隆代码**
   ```cmd
   git clone [项目地址]
   cd data-to-pdfprint\desktop-app
   ```

3. **设置环境**
   ```cmd
   # 双击运行或命令行执行
   setup_build_env.bat
   ```

### 构建可执行文件

**方式一：简化构建（推荐）**
```cmd
python simple_build.py
```

**方式二：完整构建**
```cmd
# 双击运行
build_windows.bat
```

### 构建输出
成功后将生成：
```
dist\PDFLabelGenerator\
├── PDFLabelGenerator.exe    # Windows可执行文件
├── fonts\                   # 字体文件
└── 其他依赖文件...
```

### 创建分发包
```cmd
# 进入 dist 目录并压缩
cd dist
powershell Compress-Archive -Path PDFLabelGenerator -DestinationPath PDFLabelGenerator-Windows.zip
```

## 📋 分发清单

### 完整分发包应包含
- `PDFLabelGenerator-macOS.zip` (62MB) - macOS 版本
- `PDFLabelGenerator-Windows.zip` (预估 80-120MB) - Windows 版本
- `DISTRIBUTION_GUIDE.md` (本文档) - 分发说明

## 🎯 用户使用说明

### 通用步骤
1. 下载对应操作系统的 ZIP 文件
2. 解压到任意目录
3. 运行可执行文件：
   - macOS: 双击 `PDFLabelGenerator`
   - Windows: 双击 `PDFLabelGenerator.exe`

### 功能说明
- **五步向导式界面**：简单易用的分步操作
- **三种模式支持**：常规模式、分盒模式、套盒模式
- **Excel 数据导入**：支持 .xlsx 和 .xls 格式
- **PDF 批量生成**：支持盒标和箱标批量生成
- **跨平台兼容**：Windows 和 macOS 完全兼容

### 注意事项
1. **首次运行权限**
   - macOS: 可能需要在"系统偏好设置 > 安全性与隐私"中允许
   - Windows: 可能需要在防火墙/杀毒软件中添加白名单

2. **系统要求**
   - Windows: Windows 10 及以上
   - macOS: macOS 10.14 及以上

3. **文件路径**
   - 建议解压到英文路径，避免中文字符可能导致的问题
   - 确保有足够的磁盘空间（至少 200MB 可用空间）

## 🛠️ 技术支持

如遇到问题，请检查：
1. 操作系统版本是否符合要求
2. 是否有足够的磁盘空间
3. 防火墙/杀毒软件设置
4. 文件路径是否包含特殊字符

## 📞 联系方式

如需技术支持或报告问题，请联系项目维护者。

---

💡 **提示**: 建议在分发前在目标操作系统上进行完整测试，确保所有功能正常工作。