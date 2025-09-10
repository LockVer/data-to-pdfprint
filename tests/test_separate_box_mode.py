"""
分盒模式数据生成逻辑测试

测试覆盖：
1. 数据处理逻辑 (DataProcessor._process_separate_box_mode)
2. 计算结果验证
3. 边界条件测试
4. 参数组合测试
"""

import pytest
import sys
import math
from pathlib import Path

# 添加项目根目录到路径中
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / 'src'))

from data.data_processor import DataProcessor, PackagingConfig


class TestSeparateBoxModeLogic:
    """分盒模式逻辑测试类"""
    
    def setup_method(self):
        """测试前置设置"""
        self.processor = DataProcessor()
        self.base_variables = {
            'customer_code': 'TEST001',
            'theme': 'Test Theme',
            'start_number': 'JAW00001',
            'packaging_mode': 'separate_box'
        }
    
    def create_config(self, box_quantity=100, small_box_capacity=5, large_box_capacity=3):
        """创建测试配置"""
        return PackagingConfig(
            box_quantity=box_quantity,
            set_quantity=6,  # 这个参数在分盒模式下不使用
            small_box_capacity=small_box_capacity,
            large_box_capacity=large_box_capacity,
            cards_per_box_in_set=630,
            boxes_per_set=6,
            is_overweight=False,
            sets_per_large_box=2,
            cases_per_set=1
        )
    
    def test_basic_separate_box_calculation(self):
        """测试基本的分盒模式计算"""
        # 测试用例：100盒，每小箱5盒，每大箱3小箱
        config = self.create_config(box_quantity=100, small_box_capacity=5, large_box_capacity=3)
        
        result = self.processor._process_separate_box_mode(self.base_variables, config)
        
        # 预期结果：
        # 小箱数 = ceil(100 / 5) = 20
        # 大箱数 = ceil(20 / 3) = 7
        expected_small_boxes = math.ceil(100 / 5)  # 20
        expected_large_boxes = math.ceil(expected_small_boxes / 3)  # 7
        
        assert result['box_quantity'] == 100
        assert result['small_box_quantity'] == expected_small_boxes
        assert result['large_box_quantity'] == expected_large_boxes
        assert result['boxes_per_small_box'] == 5
        assert result['small_boxes_per_large_box'] == 3
        assert result['packaging_mode'] == 'separate_box'
    
    def test_exact_division_calculation(self):
        """测试能整除的情况"""
        # 测试用例：120盒，每小箱6盒，每大箱4小箱
        config = self.create_config(box_quantity=120, small_box_capacity=6, large_box_capacity=4)
        
        result = self.processor._process_separate_box_mode(self.base_variables, config)
        
        # 预期结果：
        # 小箱数 = 120 / 6 = 20
        # 大箱数 = 20 / 4 = 5
        assert result['small_box_quantity'] == 20
        assert result['large_box_quantity'] == 5
    
    def test_edge_case_single_box_per_container(self):
        """测试边界情况：每小箱1盒，每大箱1小箱"""
        config = self.create_config(box_quantity=10, small_box_capacity=1, large_box_capacity=1)
        
        result = self.processor._process_separate_box_mode(self.base_variables, config)
        
        # 预期结果：小箱数 = 大箱数 = 盒数 = 10
        assert result['small_box_quantity'] == 10
        assert result['large_box_quantity'] == 10
        assert result['boxes_per_small_box'] == 1
        assert result['small_boxes_per_large_box'] == 1
    
    def test_edge_case_zero_capacity(self):
        """测试边界情况：容量为0"""
        config = self.create_config(box_quantity=50, small_box_capacity=0, large_box_capacity=0)
        
        result = self.processor._process_separate_box_mode(self.base_variables, config)
        
        # 预期结果：当容量为0时，应该使用原始盒数量
        assert result['small_box_quantity'] == 50  # 当small_box_capacity为0时，小箱数等于盒数
        assert result['large_box_quantity'] == 50  # 当large_box_capacity为0时，大箱数等于小箱数
    
    def test_large_numbers(self):
        """测试大数值情况"""
        config = self.create_config(box_quantity=10000, small_box_capacity=50, large_box_capacity=20)
        
        result = self.processor._process_separate_box_mode(self.base_variables, config)
        
        expected_small_boxes = math.ceil(10000 / 50)  # 200
        expected_large_boxes = math.ceil(200 / 20)    # 10
        
        assert result['small_box_quantity'] == expected_small_boxes
        assert result['large_box_quantity'] == expected_large_boxes
    
    def test_label_specifications(self):
        """测试标签规格计算"""
        config = self.create_config(box_quantity=75, small_box_capacity=8, large_box_capacity=3)
        
        result = self.processor._process_separate_box_mode(self.base_variables, config)
        
        # 检查标签规格
        label_specs = result['label_specifications']
        assert label_specs['box_labels'] == 75
        assert label_specs['small_box_labels'] == result['small_box_quantity']
        assert label_specs['large_box_labels'] == result['large_box_quantity']
    
    def test_result_data_structure(self):
        """测试返回数据结构完整性"""
        config = self.create_config()
        
        result = self.processor._process_separate_box_mode(self.base_variables, config)
        
        # 检查必要的字段是否存在
        required_fields = [
            'customer_code', 'theme', 'start_number', 'packaging_mode',
            'box_quantity', 'small_box_quantity', 'large_box_quantity',
            'boxes_per_small_box', 'small_boxes_per_large_box', 'label_specifications'
        ]
        
        for field in required_fields:
            assert field in result, f"Missing field: {field}"
        
        # 检查 label_specifications 的结构
        assert 'box_labels' in result['label_specifications']
        assert 'small_box_labels' in result['label_specifications']
        assert 'large_box_labels' in result['label_specifications']
    
    def test_base_data_inheritance(self):
        """测试基础数据传递"""
        custom_base_data = {
            'customer_code': 'CUSTOM123',
            'theme': 'Custom Theme',
            'start_number': 'ABC00999',
            'packaging_mode': 'separate_box'
        }
        config = self.create_config()
        
        result = self.processor._process_separate_box_mode(custom_base_data, config)
        
        # 验证基础数据正确传递
        assert result['customer_code'] == 'CUSTOM123'
        assert result['theme'] == 'Custom Theme'
        assert result['start_number'] == 'ABC00999'
        assert result['packaging_mode'] == 'separate_box'


class TestSeparateBoxModeParameterCombinations:
    """分盒模式参数组合测试"""
    
    def setup_method(self):
        """测试前置设置"""
        self.processor = DataProcessor()
        self.base_variables = {
            'customer_code': 'TEST001',
            'theme': 'Test Theme',
            'start_number': 'JAW00001',
            'packaging_mode': 'separate_box'
        }
    
    @pytest.mark.parametrize("box_qty,small_cap,large_cap,expected_small,expected_large", [
        # (盒数量, 小箱容量, 大箱容量, 期望小箱数, 期望大箱数)
        (50, 10, 5, 5, 1),      # 正好整除
        (51, 10, 5, 6, 2),      # 小箱需要向上取整
        (50, 10, 3, 5, 2),      # 大箱需要向上取整
        (37, 7, 4, 6, 2),       # 两级都需要向上取整
        (100, 1, 1, 100, 100),  # 每箱1盒
        (1000, 100, 10, 10, 1), # 大数值正好整除
    ])
    def test_parameter_combinations(self, box_qty, small_cap, large_cap, expected_small, expected_large):
        """测试各种参数组合"""
        config = PackagingConfig(
            box_quantity=box_qty,
            set_quantity=6,
            small_box_capacity=small_cap,
            large_box_capacity=large_cap,
            cards_per_box_in_set=630,
            boxes_per_set=6,
            is_overweight=False,
            sets_per_large_box=2,
            cases_per_set=1
        )
        
        result = self.processor._process_separate_box_mode(self.base_variables, config)
        
        assert result['small_box_quantity'] == expected_small, \
            f"Expected {expected_small} small boxes, got {result['small_box_quantity']}"
        assert result['large_box_quantity'] == expected_large, \
            f"Expected {expected_large} large boxes, got {result['large_box_quantity']}"


class TestSeparateBoxModeIntegration:
    """分盒模式集成测试"""
    
    def setup_method(self):
        """测试前置设置"""
        self.processor = DataProcessor()
    
    def test_process_for_packaging_mode_separate_box(self):
        """测试通过 process_for_packaging_mode 调用分盒模式"""
        variables = {
            'customer_code': 'INT001',
            'theme': 'Integration Test',
            'start_number': 'INT00001'
        }
        
        config = PackagingConfig(
            box_quantity=200,
            set_quantity=6,
            small_box_capacity=25,
            large_box_capacity=8,
            cards_per_box_in_set=630,
            boxes_per_set=6,
            is_overweight=False,
            sets_per_large_box=2,
            cases_per_set=1
        )
        
        result = self.processor.process_for_packaging_mode(variables, config, 'separate_box')
        
        # 预期结果计算
        expected_small_boxes = math.ceil(200 / 25)  # 8
        expected_large_boxes = math.ceil(8 / 8)     # 1
        
        assert result['packaging_mode'] == 'separate_box'
        assert result['small_box_quantity'] == expected_small_boxes
        assert result['large_box_quantity'] == expected_large_boxes
    
    def test_calculate_label_ranges_for_separate_box(self):
        """测试分盒模式的标签编号范围计算"""
        processed_data = {
            'packaging_mode': 'separate_box',
            'start_number': 'JAW00100',
            'box_quantity': 15
        }
        
        ranges = self.processor.calculate_label_ranges(processed_data)
        
        # 验证盒号范围
        expected_box_numbers = [f"JAW{100 + i:05d}" for i in range(15)]
        assert ranges['box_numbers'] == expected_box_numbers


if __name__ == '__main__':
    # 运行测试的简单方法
    pytest.main([__file__, '-v'])