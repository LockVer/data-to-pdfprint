# 测试套件

Data to PDF Print项目的完整测试套件，采用分层组织结构确保测试的清晰性和可维护性。

## 📁 目录结构

```
tests/
├── unit/           # 单元测试 - 核心功能快速验证
├── integration/    # 集成测试 - 复杂场景和模块协调
├── dev/           # 开发辅助 - 快速验证和调试工具
├── docs/          # 测试文档 - 详细说明和指导
├── quick_test.py  # 通用快速测试入口
└── README.md      # 本文件 - 测试导航
```

## 🚀 快速开始

### 运行所有核心测试
```bash
# 运行所有单元测试 (推荐日常使用)
python tests/unit/test_quantity_quick.py
python tests/unit/test_serial_quick.py  
python tests/unit/test_carton_logic_fixed.py

# 或使用通用入口
python tests/quick_test.py
```

### 运行特定类型测试
```bash
# 单元测试 - 快速核心功能验证
cd tests/unit/
python test_quantity_quick.py
python test_serial_quick.py
python test_carton_logic_fixed.py

# 集成测试 - 复杂场景验证  
cd tests/integration/
python test_serial_logic_comprehensive.py
python test_carton_logic.py

# 开发验证 - 调试工具
cd tests/dev/
python test_quantity_quick_validation.py
```

## 📊 测试覆盖

### 单元测试 (`unit/`)
- **量计算测试** (`test_quantity_quick.py`) - 验证quantity计算逻辑的正确性
- **序列号测试** (`test_serial_quick.py`) - 验证serial生成逻辑  
- **箱号测试** (`test_carton_logic_fixed.py`) - 验证carton number计算

### 集成测试 (`integration/`)  
- **序列号综合测试** (`test_serial_logic_comprehensive.py`) - 复杂场景的serial逻辑
- **箱号综合测试** (`test_carton_logic.py`) - 箱号计算的集成验证

### 开发辅助 (`dev/`)
- **快速验证** (`test_quantity_quick_validation.py`) - 开发时的轻量级验证工具

## 🎯 测试理念

### 分层设计
- **Unit**: 简洁高效，专注核心功能，快速反馈
- **Integration**: 复杂场景，模块协调，深度验证  
- **Dev**: 开发辅助，调试友好，快速迭代

### 风格统一
所有测试采用一致的函数式风格：
- 简洁的测试函数
- 清晰的输出格式
- 易于理解和维护

### 质量保证
- 功能正确性验证
- 边界情况处理
- 真实场景覆盖
- 性能基准检查

## 📖 详细文档

- [Quantity测试说明](docs/README_QUANTITY_TESTS.md) - quantity计算测试的详细说明
- [综合测试更新](docs/COMPREHENSIVE_TEST_UPDATE.md) - 测试套件的更新历史
- [测试指南](docs/TEST_README.md) - 测试的详细指导

## 🔧 开发指南

### 添加新测试
1. **单元测试**: 添加到 `unit/` 目录，用于快速核心功能验证
2. **集成测试**: 添加到 `integration/` 目录，用于复杂场景验证
3. **开发工具**: 添加到 `dev/` 目录，用于调试和验证

### 测试命名约定
- 单元测试: `test_[功能]_quick.py`
- 集成测试: `test_[功能]_logic_comprehensive.py` 或 `test_[功能]_logic.py`
- 开发工具: `test_[功能]_validation.py`

### 风格指南
- 使用简洁的函数式测试
- 保持与现有测试风格一致
- 提供清晰的测试输出
- 包含必要的边界情况测试

---

💡 **提示**: 日常开发建议主要使用 `unit/` 目录下的测试，它们执行快速且覆盖核心功能。复杂场景验证时再使用 `integration/` 测试。