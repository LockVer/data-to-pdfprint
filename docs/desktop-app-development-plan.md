# PDF标签生成系统 - 桌面应用开发方案

## 1. 技术选择

### 推荐技术栈
- **Python 3.8+**: 复用现有后端代码
- **Tkinter**: 内置GUI库，跨平台支持
- **cx_Freeze/PyInstaller**: 打包成可执行文件
- **Pillow**: 图像处理和图标支持

### 技术选择对比

| 方案 | 优势 | 劣势 | 适用性 |
|------|------|------|--------|
| **Python + Tkinter** | 轻量级、复用现有代码、跨平台 | 界面相对简单 | ✅ 推荐 |
| **Python + PyQt** | 界面美观、功能丰富 | 许可证费用、体积大 | 备选 |
| **Electron + Vue** | 现代界面、前端技术栈 | 体积大、性能开销 | 不推荐 |

## 2. 应用架构设计

### 项目结构
```
desktop-app/
├── main.py                 # 主程序入口
├── gui/                   # GUI模块
│   ├── __init__.py
│   ├── main_window.py     # 主窗口
│   ├── components/        # UI组件
│   │   ├── __init__.py
│   │   ├── file_upload.py    # 文件上传组件
│   │   ├── mode_selector.py  # 模式选择组件
│   │   ├── param_input.py    # 参数输入组件
│   │   ├── template_selector.py # 模板选择组件
│   │   └── progress_bar.py   # 进度条组件
│   ├── dialogs/           # 对话框
│   │   ├── __init__.py
│   │   ├── error_dialog.py
│   │   └── success_dialog.py
│   └── utils/             # GUI工具
│       ├── __init__.py
│       ├── validators.py  # 输入验证
│       └── formatters.py  # 格式化工具
├── core/                  # 核心业务逻辑
│   ├── __init__.py
│   ├── pdf_generator.py   # PDF生成器
│   ├── excel_processor.py # Excel处理器
│   └── config_manager.py  # 配置管理
├── assets/                # 资源文件
│   ├── icons/            # 图标文件
│   │   ├── app.ico
│   │   ├── file.png
│   │   └── generate.png
│   ├── fonts/            # 字体文件
│   └── styles/           # 样式文件
├── tests/                 # 测试文件
│   ├── __init__.py
│   ├── test_gui.py
│   └── test_core.py
├── requirements.txt       # 依赖包列表
├── setup.py              # 打包配置
└── README.md             # 说明文档
```

### 核心类设计

```python
# 主应用类
class PDFLabelGeneratorApp:
    def __init__(self):
        self.root = tk.Tk()
        self.current_step = 0
        self.app_data = AppData()
        self.setup_ui()
    
    def setup_ui(self):
        # 设置主窗口
        # 创建步骤组件
        # 绑定事件处理
        pass

# 应用数据模型
class AppData:
    def __init__(self):
        self.excel_file_path = None
        self.excel_data = None
        self.package_mode = None
        self.package_params = {}
        self.label_template = None

# 步骤基类
class BaseStep:
    def __init__(self, parent, app_data):
        self.parent = parent
        self.app_data = app_data
        self.frame = ttk.Frame(parent)
        self.setup_ui()
    
    def setup_ui(self):
        raise NotImplementedError
    
    def validate(self):
        raise NotImplementedError
```

## 3. 界面设计规范

### 窗口设计
- **窗口尺寸**: 800x600px (最小), 1000x700px (推荐)
- **窗口标题**: "PDF标签生成器 v1.0"
- **图标**: 自定义应用图标
- **可调整大小**: 是

### 布局设计
```python
# 主窗口布局
┌─────────────────────────────────────────────────┐
│ [标题栏] PDF标签生成器                            │
├─────────────────────────────────────────────────┤
│ [进度条] ●─○─○─○─○  步骤1/5                      │
├─────────────────────────────────────────────────┤
│                                                 │
│                [内容区域]                        │
│                                                 │
│                                                 │
│                                                 │
├─────────────────────────────────────────────────┤
│              [上一步] [下一步/生成]              │
└─────────────────────────────────────────────────┘
```

### 配色方案
- **主色调**: #2196F3 (蓝色)
- **辅助色**: #FFC107 (橙色)
- **成功色**: #4CAF50 (绿色)
- **错误色**: #F44336 (红色)
- **背景色**: #FAFAFA (浅灰)

## 4. 实现步骤规划

### 第一阶段: 基础框架 (1-2天)
- [x] 创建项目结构
- [ ] 实现主窗口框架
- [ ] 创建步骤切换机制
- [ ] 实现进度条组件

### 第二阶段: 核心组件 (3-4天)
- [ ] 文件上传组件
- [ ] 模式选择组件  
- [ ] 参数输入组件
- [ ] 模板选择组件
- [ ] 结果展示组件

### 第三阶段: 业务逻辑 (2-3天)
- [ ] 集成Excel处理代码
- [ ] 集成PDF生成代码
- [ ] 数据验证逻辑
- [ ] 错误处理机制

### 第四阶段: 优化完善 (2-3天)
- [ ] 界面美化
- [ ] 交互优化
- [ ] 性能优化
- [ ] 单元测试

### 第五阶段: 打包发布 (1天)
- [ ] 打包配置
- [ ] 多平台测试
- [ ] 安装程序制作
- [ ] 文档编写

## 5. 技术实现细节

### GUI组件实现

#### 文件上传组件
```python
class FileUploadStep(BaseStep):
    def setup_ui(self):
        # 文件选择按钮
        self.select_btn = ttk.Button(
            self.frame, 
            text="选择Excel文件", 
            command=self.select_file
        )
        
        # 文件路径显示
        self.file_path_var = tk.StringVar()
        self.file_path_label = ttk.Label(
            self.frame, 
            textvariable=self.file_path_var
        )
        
        # 文件信息预览
        self.preview_text = tk.Text(
            self.frame, 
            height=10, 
            state='disabled'
        )
```

#### 参数输入组件
```python
class ParamInputStep(BaseStep):
    def setup_ui(self):
        # 根据包装模式动态显示参数输入
        self.create_param_inputs()
    
    def create_param_inputs(self):
        mode = self.app_data.package_mode
        
        # 每盒张数 (所有模式必需)
        self.sheets_per_box = tk.IntVar()
        
        if mode in ['separate', 'set']:
            # 每小箱盒数
            self.boxes_per_small_case = tk.IntVar()
            # 每大箱小箱数
            self.small_cases_per_large_case = tk.IntVar()
```

### 数据处理

#### Excel数据提取
```python
class ExcelProcessor:
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = {}
    
    def extract_data(self):
        # 复用现有的ExcelReader类
        from data.excel_reader import ExcelReader
        reader = ExcelReader(self.file_path)
        self.data = reader.extract_template_variables()
        return self.data
```

#### PDF生成接口
```python
class PDFGeneratorService:
    def __init__(self):
        pass
    
    def generate_labels(self, excel_data, config):
        # 生成盒标
        box_pdf = self.generate_box_labels(excel_data, config)
        
        # 生成箱标  
        case_pdf = self.generate_case_labels(excel_data, config)
        
        return box_pdf, case_pdf
```

## 6. 打包部署方案

### PyInstaller配置
```python
# setup.py
from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": ["tkinter", "PIL", "reportlab", "pandas"],
    "include_files": [
        ("assets/", "assets/"),
        ("src/fonts/", "fonts/")
    ],
    "excludes": ["test"]
}

setup(
    name="PDF标签生成器",
    version="1.0",
    description="游戏卡片标签生成工具",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base="Win32GUI")]
)
```

### 平台特定配置

#### Windows
```bash
# 生成Windows可执行文件
python setup.py build
# 或使用PyInstaller
pyinstaller --onefile --windowed main.py
```

#### macOS
```bash
# 生成macOS应用包
python setup.py bdist_mac
# 或使用PyInstaller
pyinstaller --onefile --windowed --add-data "assets:assets" main.py
```

#### Linux
```bash
# 生成Linux可执行文件
python setup.py build
```

## 7. 测试策略

### 单元测试
- GUI组件测试
- 数据处理测试
- PDF生成测试

### 集成测试  
- 完整流程测试
- 多种Excel文件测试
- 错误场景测试

### 用户测试
- 易用性测试
- 跨平台兼容性测试
- 性能压力测试

## 8. 维护和更新

### 版本控制
- 语义化版本号 (v1.0.0)
- 自动更新检测机制
- 配置文件向后兼容

### 日志和监控
- 应用日志记录
- 错误报告机制
- 使用情况统计

### 文档和支持
- 用户使用手册
- 常见问题解答
- 技术支持联系方式