# Data to PDF Print

一个用于读取Excel数据并生成PDF标签的Python GUI/CLI工具。支持多种标签类型（常规盒标、分盒标、嵌套盒标），可以根据Excel数据快速生成专业的PDF标签文件。

## 项目结构

```
data-to-pdfprint/
├── scripts/                     # 构建脚本目录
│   ├── build_gui.py             # macOS GUI构建脚本
│   └── build_windows.py         # Windows GUI构建脚本
├── src/                         # 源代码目录
│   ├── gui_app.py               # GUI主程序 (直接运行此文件进行调试)
│   ├── fonts/                   # 字体文件目录
│   │   ├── msyh.ttf             # 微软雅黑字体
│   │   └── msyhbd.ttc           # 微软雅黑粗体字体
│   ├── data/                    # 数据处理模块
│   ├── pdf/                     # PDF生成模块
│   │   ├── regular_box/         # 常规盒标
│   │   ├── split_box/           # 分盒标
│   │   └── nested_box/          # 嵌套盒标
│   └── utils/                   # 工具函数模块
├── requirements.txt             # Python依赖包列表
├── .gitignore                   # Git忽略文件配置
├── CLAUDE.md                    # 开发指导文档
└── README.md                    # 项目说明文档
```

## 快速开始

### 方法一：直接运行源代码（推荐用于开发调试）

#### 1. 环境准备
确保系统已安装 Python 3.8 或更高版本：
```bash
python3 --version
```

#### 2. 克隆或下载项目
```bash
git clone https://github.com/yourusername/data-to-pdfprint.git
cd data-to-pdfprint
```

#### 3. 创建虚拟环境
```bash
# 创建虚拟环境
python3 -m venv venv

# 激活虚拟环境
# macOS/Linux:
source venv/bin/activate
# Windows:
venv\Scripts\activate
```

#### 4. 安装依赖
```bash
pip install -r requirements.txt
```

#### 5. 直接运行GUI程序
```bash
# 进入源码目录运行GUI程序
python3 src/gui_app.py
```

### 方法二：构建独立可执行文件

#### macOS版本构建
```bash
# 1. 确保已安装依赖
pip install -r requirements.txt

# 2. 运行macOS构建脚本
python3 scripts/build_gui.py

# 3. 运行生成的应用
./dist/DataToPDF_GUI
```

#### Windows版本构建
```bash
# 1. 确保已安装依赖（最好在Windows系统上运行）
pip install -r requirements.txt

# 2. 运行Windows构建脚本
python scripts/build_windows.py

# 3. 运行生成的程序
dist/DataToPDF_GUI.exe
```

### Excel文件格式要求

程序支持标准的Excel格式(.xlsx, .xls)，具体的数据格式要求请参考各标签类型的说明：

- **常规盒标**：需要包含产品名称、规格、批次等基本信息
- **分盒标**：除基本信息外，还需包含分装相关数据
- **嵌套盒标**：需要包含层级结构信息

## 依赖包说明

```
click>=8.1.0          # 命令行框架
reportlab>=3.6.0      # PDF生成库
pandas>=1.5.0         # 数据处理库
openpyxl>=3.1.0       # Excel文件读取库
pytest>=7.0.0         # 测试框架
black>=22.0.0         # 代码格式化工具
flake8>=5.0.0         # 代码质量检查工具
pyinstaller>=5.0.0    # 程序打包工具
```

## 开发说明

### 项目架构

- **数据层** (`src/data/`): 负责Excel数据的读取和处理
- **PDF生成层** (`src/pdf/`): 按不同标签类型生成PDF文件
  - `regular_box/`: 常规盒标相关代码
  - `split_box/`: 分盒标相关代码  
  - `nested_box/`: 嵌套盒标相关代码
- **工具层** (`src/utils/`): 提供通用的工具函数
- **GUI层** (`src/gui_app.py`): 图形用户界面

### 调试和开发

1. **直接运行GUI进行调试**：
   ```bash
   python3 src/gui_app.py
   ```

2. **代码格式化**：
   ```bash
   black src/
   ```

3. **代码质量检查**：
   ```bash
   flake8 src/
   ```

4. **运行测试**（如果有测试文件）：
   ```bash
   pytest tests/
   ```

### 构建脚本说明

- `scripts/build_gui.py`: macOS版本构建脚本，生成适用于macOS的应用程序
- `scripts/build_windows.py`: Windows版本构建脚本，生成适用于Windows的exe文件

构建后的文件包含完整的运行环境和字体文件，可以在目标系统上独立运行，无需安装Python。

## 字体支持

项目内置了微软雅黑字体文件：
- `src/fonts/msyh.ttf`: 微软雅黑常规字体
- `src/fonts/msyhbd.ttc`: 微软雅黑粗体字体

这些字体会被自动打包到可执行文件中，确保在任何系统上都能正确显示中文内容。

## 分发说明

### 文件大小
- macOS版本：约70MB（包含字体文件）
- Windows版本：约110MB（包含字体文件）

### 系统兼容性
- **macOS**: macOS 10.14+ (支持Intel和Apple Silicon)
- **Windows**: Windows 7/8/10/11 (32/64位)