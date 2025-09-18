# 测试目录说明

本目录包含 Data-to-PDFPrint 项目的测试套件，专门用于验证 Carton Number 计算逻辑。

## 🗂️ 目录结构

```
tests/
├── README.md                           # 本说明文件
├── test_carton_logic.py                # 完整测试套件
├── quick_test.py                       # 快速测试工具
├── TEST_README.md                      # 详细测试说明文档
├── COMPREHENSIVE_TEST_UPDATE.md        # 测试实现总结
├── test_results_YYYYMMDD_HHMMSS.json  # 测试数据文件 (自动生成)
├── test_report_YYYYMMDD_HHMMSS.md     # 测试报告 (自动生成)
└── quick_test_results_YYYYMMDD_HHMMSS.md # 快速测试结果 (自动生成)
```

## 🚀 快速开始

### 运行完整测试套件
```bash
cd tests
python test_carton_logic.py
```

### 运行快速测试
```bash
cd tests
python quick_test.py
```

## 📋 测试覆盖

### 二级模式测试 (5个)
- 一套分多个大箱：多级编号 (1-1, 1-2, 2-1...)
- 一套分一个大箱：单级编号 (1, 2, 3...)  
- 多套分一个大箱：范围编号 (1-2, 3-4...)

### 三级模式测试 (6个)
- 小箱标：多级/单级/不生成
- 大箱标：多级/单级/范围编号

### 边界和性能测试 (4个)
- 最小值、整除、余数、大数量

## 📄 输出文件

所有测试结果文件都带有时间戳，包括：
- **JSON 数据文件**: 完整的测试数据，适合程序化分析
- **Markdown 报告**: 格式化的测试报告，适合人工阅读
- **快速测试结果**: 核心场景的验证结果

这些文件已在 `.gitignore` 中配置为忽略，不会被提交到版本控制。

## 🔧 自定义测试

详细的测试说明和自定义指南请参考：
- `TEST_README.md` - 完整的测试框架使用说明
- `COMPREHENSIVE_TEST_UPDATE.md` - 测试实现技术总结

## ⚠️ 注意事项

- 测试文件需要在项目根目录或 tests 目录运行
- 确保已安装项目依赖：`pip install -r requirements.txt`
- 测试过程中会生成详细的调试输出，便于问题定位