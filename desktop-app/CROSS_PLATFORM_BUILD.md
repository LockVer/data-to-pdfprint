# 跨平台构建方案

## 🎯 问题说明
PyInstaller 无法跨平台构建，在 macOS 上无法直接生成 Windows 可执行文件。

## 🔄 解决方案

### 方案一：虚拟机构建 (推荐)
1. **在 macOS 上运行 Windows 虚拟机**
   - 使用 Parallels Desktop、VMware Fusion 或 VirtualBox
   - 安装 Windows 10/11 虚拟机
   - 在虚拟机中执行构建

2. **构建步骤**
   ```cmd
   # 在 Windows 虚拟机中
   git clone [项目地址]
   cd data-to-pdfprint\desktop-app
   setup_build_env.bat
   python simple_build.py
   ```

### 方案二：GitHub Actions (自动化)
创建 GitHub Actions 工作流自动构建多平台版本：

```yaml
# .github/workflows/build.yml
name: Build Multi-Platform
on: [push, release]

jobs:
  build:
    runs-on: ${{ matrix.os }}
    strategy:
      matrix:
        os: [windows-latest, macos-latest]
    
    steps:
    - uses: actions/checkout@v3
    - name: Set up Python
      uses: actions/setup-python@v4
      with:
        python-version: '3.11'
    
    - name: Install dependencies
      run: |
        cd desktop-app
        pip install -r requirements.txt
    
    - name: Build application
      run: |
        cd desktop-app  
        python simple_build.py
    
    - name: Create distribution
      run: |
        cd desktop-app/dist
        # Windows
        if [ "$RUNNER_OS" == "Windows" ]; then
          powershell Compress-Archive -Path PDFLabelGenerator -DestinationPath PDFLabelGenerator-Windows.zip
        fi
        # macOS
        if [ "$RUNNER_OS" == "macOS" ]; then
          zip -r PDFLabelGenerator-macOS.zip PDFLabelGenerator/
        fi
    
    - name: Upload artifacts
      uses: actions/upload-artifact@v3
      with:
        name: PDFLabelGenerator-${{ matrix.os }}
        path: desktop-app/dist/PDFLabelGenerator-*.zip
```

### 方案三：Docker 容器化
```dockerfile
# Windows 构建容器
FROM python:3.11-windowsservercore

WORKDIR /app
COPY . .
RUN pip install -r requirements.txt
RUN python simple_build.py

# 输出构建结果
VOLUME ["/app/dist"]
```

### 方案四：云构建服务
- **GitHub Codespaces**: 在线开发环境
- **AWS EC2**: Windows 实例构建
- **Azure VM**: Windows 虚拟机服务

## 🚀 快速实施建议

### 最经济方案：虚拟机
1. 下载 VirtualBox (免费)
2. 安装 Windows 10/11 VM
3. 分配 4GB+ 内存，20GB+ 磁盘
4. 在 VM 中克隆代码并构建

### 最自动化方案：GitHub Actions
1. 将代码推送到 GitHub
2. 添加上述工作流文件
3. 每次提交自动构建两个平台版本
4. 从 Actions 页面下载构建结果

## 📋 当前状态
- ✅ macOS 版本已完成构建
- ⏳ Windows 版本需要在 Windows 环境中构建
- 📁 所有构建脚本和配置已准备就绪

## 💡 建议
推荐使用 **GitHub Actions** 方案，一次设置后可以自动化构建两个平台版本，无需维护本地虚拟机。