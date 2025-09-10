# 分盒模式数据生成逻辑测试

本测试套件专门用于测试分盒模式的数据生成逻辑，避免通过UI操作进行繁琐的手工测试。

## 测试覆盖范围

### 1. 数据处理逻辑测试 (`test_separate_box_mode.py`)
- **基础计算逻辑**: 测试小箱数量和大箱数量的计算
- **边界条件测试**: 测试除零、向上取整等边界情况
- **参数组合测试**: 测试各种参数组合的计算正确性
- **集成测试**: 测试完整的数据处理流程

### 2. 标签生成逻辑测试 (`test_separate_box_labels.py`)
- **标签数量计算**: 验证小箱标和大箱标的quantity计算规律
- **序列号生成**: 验证serial字段的生成规律
- **箱号生成**: 验证carton_no字段的生成规律
- **标签模拟**: 完整模拟标签生成过程

## 快速开始

### 安装依赖
```bash
pip install pytest
```

### 运行测试

#### 1. 使用测试运行器（推荐）
```bash
# 运行演示示例
python test_separate_box.py --demo

# 运行基础逻辑测试
python test_separate_box.py --basic

# 运行标签生成测试
python test_separate_box.py --labels

# 运行完整测试套件
python test_separate_box.py --all
```

#### 2. 直接使用pytest
```bash
# 运行所有测试
pytest tests/test_separate_box_mode.py tests/test_separate_box_labels.py -v

# 只运行数据处理测试
pytest tests/test_separate_box_mode.py -v

# 只运行标签生成测试
pytest tests/test_separate_box_labels.py -v
```

## 测试场景说明

### 基础计算测试场景

| 场景 | 盒数量 | 每小箱盒数 | 每大箱小箱数 | 预期小箱数 | 预期大箱数 |
|------|--------|------------|--------------|------------|------------|
| 正好整除 | 120 | 6 | 4 | 20 | 5 |
| 需要向上取整 | 100 | 7 | 3 | 15 | 5 |
| 边界条件 | 10 | 1 | 1 | 10 | 10 |
| 大数值 | 10000 | 50 | 20 | 200 | 10 |

### 标签生成测试场景

以27盒卡片为例，每盒1000张，每小箱6盒，每大箱2小箱：

**计算结果**:
- 需要小箱: 5个 (6,6,6,6,3盒)
- 需要大箱: 3个 (2,2,1小箱)

**小箱标示例**:
```
小箱 1: 6000PCS | LAB00100-LAB00105 | 箱号:01-01 | 含6盒
小箱 2: 6000PCS | LAB00106-LAB00111 | 箱号:01-02 | 含6盒
小箱 3: 6000PCS | LAB00112-LAB00117 | 箱号:02-01 | 含6盒
小箱 4: 6000PCS | LAB00118-LAB00123 | 箱号:02-02 | 含6盒
小箱 5: 3000PCS | LAB00124-LAB00126 | 箱号:03-01 | 含3盒
```

**大箱标示例**:
```
大箱 1: 12000PCS | LAB00100-LAB00111 | 箱号:01 | 含2小箱
大箱 2: 12000PCS | LAB00112-LAB00123 | 箱号:02 | 含2小箱
大箱 3:  6000PCS | LAB00124-LAB00126 | 箱号:03 | 含1小箱
```

## 测试规律验证

### 小箱标生成规律
根据文档定义：
- **quantity**: 盒张数 × 每小箱盒数
- **serial**: 该小箱包含的盒子序列号范围
- **carton_no**: 大箱编号-小箱编号 (如: 01-01, 01-02)

### 大箱标生成规律
根据文档定义：
- **quantity**: 大箱内小箱数量 × 盒张数 × 每小箱盒数
- **serial**: 该大箱包含的所有盒子序列号范围
- **carton_no**: 大箱编号 (如: 01, 02, 03)

## 自定义测试

### 添加新的测试场景
在 `TestSeparateBoxModeParameterCombinations` 类中添加参数化测试：

```python
@pytest.mark.parametrize("box_qty,small_cap,large_cap,expected_small,expected_large", [
    # 添加新的测试案例
    (你的盒数量, 小箱容量, 大箱容量, 期望小箱数, 期望大箱数),
])
def test_your_scenario(self, box_qty, small_cap, large_cap, expected_small, expected_large):
    # 测试逻辑
```

### 验证实际业务数据
创建基于实际业务数据的测试：

```python
def test_real_business_case(self):
    """测试实际业务案例"""
    config = PackagingConfig(
        box_quantity=你的实际盒数,
        small_box_capacity=你的实际小箱容量,
        large_box_capacity=你的实际大箱容量,
        # ... 其他参数
    )
    
    result = self.processor._process_separate_box_mode(self.base_variables, config)
    
    # 验证结果是否符合预期
    assert result['small_box_quantity'] == 你的预期小箱数
    assert result['large_box_quantity'] == 你的预期大箱数
```

## 性能测试

测试套件包含大数值测试（如10000盒），用于验证算法在处理大量数据时的性能和正确性。

## 故障排查

### 常见问题

1. **ImportError**: 确保项目路径正确，tests目录中的测试文件可以找到src目录下的模块
2. **计算结果不符合预期**: 检查向上取整逻辑，特别是在除法不能整除的情况
3. **序列号格式错误**: 验证起始编号的解析逻辑和前缀提取

### 调试技巧

在测试中添加调试输出：
```python
print(f"Small boxes: {result['small_box_quantity']}")
print(f"Large boxes: {result['large_box_quantity']}")
```

使用pytest的详细输出：
```bash
pytest tests/test_separate_box_mode.py -v -s
```

## 扩展测试

这个测试框架可以作为其他模式（如套盒模式、常规模式）测试的模板，只需要：

1. 复制测试文件结构
2. 修改对应的处理方法调用
3. 调整测试场景和预期结果
4. 更新文档规律验证

通过这种方式，可以为所有包装模式建立完整的自动化测试覆盖。