# PDF标签生成器 - 桌面应用

一个用于生成游戏卡片标签的跨平台桌面应用程序。

## 功能特点

- 📁 **Excel数据导入**: 支持读取Excel文件中的游戏数据
- 🏷️ **多种包装模式**: 常规模式、分盒模式、套盒模式
- 🎨 **双模板选择**: Regular模板和Game模板
- 📦 **自动配套生成**: 同时生成盒标和箱标PDF
- 🖥️ **跨平台支持**: Windows、macOS、Linux
- 🎯 **直观界面**: 步骤式向导，操作简单

## 系统要求

- Python 3.8或更高版本
- 支持tkinter的Python环境（大多数Python安装包都包含）

## 安装和运行

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 运行应用

```bash
python run.py
```

或者直接运行：

```bash
python main.py
```

## 使用说明

### 步骤1: 上传Excel文件
- 选择包含游戏数据的Excel文件（.xlsx或.xls格式）
- 确保Excel文件包含必要字段：
  - A3: 客户名称编码
  - B3: 主题
  - B10: 开始号
  - F3: 总张数

### 步骤2: 选择包装模式
- **常规模式（单级）**: 单级序列号，适合简单包装
- **分盒模式（多级）**: 多级序列号，支持分盒管理
- **套盒模式（多级）**: 多级序列号，支持套装管理

### 步骤3: 输入包装参数
- **每盒张数**: 每个盒子包含的卡片数量
- **每小箱盒数**: 每个小箱包含的盒子数量（分盒/套盒模式）
- **每大箱小箱数**: 每个大箱包含的小箱数量（分盒/套盒模式）

### 步骤4: 选择标签模板
- **Regular模板**: 传统格式（客户编码+主题+序列号）
- **Game模板**: 游戏格式（Game title + Ticket count + Serial）

### 步骤5: 生成PDF
- 点击"生成PDF"按钮
- 等待生成完成
- 查看生成的文件

## 输出格式

生成的文件结构：
```
{客户名称}+{订单名称}+标签/
├── {客户名称}+{订单名称}+盒标.pdf
└── {客户名称}+{订单名称}+箱标.pdf
```

- **格式**: PDF/X（适合CMYK打印）
- **尺寸**: 90mm × 50mm
- **分辨率**: 300 DPI

## 项目结构

```
desktop-app/
├── main.py                 # 主程序入口
├── run.py                  # 启动脚本（推荐使用）
├── requirements.txt        # 依赖包列表
├── gui/                   # GUI模块
│   ├── main_window.py     # 主窗口
│   └── components/        # UI组件
│       ├── base_step.py
│       ├── file_upload.py
│       ├── mode_selector.py
│       ├── param_input.py
│       ├── template_selector.py
│       ├── generate_result.py
│       └── progress_bar.py
├── core/                  # 核心业务逻辑
│   └── app_data.py        # 数据模型
└── README.md             # 说明文档
```

## 常见问题

### Q: 应用启动失败怎么办？
A: 请检查：
1. Python版本是否为3.8或更高
2. 是否安装了所有依赖包
3. 是否在正确的目录中运行

### Q: Excel文件读取失败？
A: 请确保：
1. 文件格式为.xlsx或.xls
2. 文件未被其他程序占用
3. Excel中包含必要的数据字段

### Q: PDF生成失败？
A: 可能原因：
1. 输出目录权限不足
2. 磁盘空间不足
3. Excel数据格式错误

## 开发信息

- **技术栈**: Python + tkinter
- **PDF生成**: ReportLab
- **Excel处理**: pandas + openpyxl
- **支持平台**: Windows, macOS, Linux

## 许可证

本项目采用MIT许可证。