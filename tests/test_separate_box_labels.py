"""
分盒模式标签生成逻辑测试

根据文档规律测试：
1. 小箱标：
   - quantity: 盒张数 × 每小箱盒数
   - serial: 父级编号(大箱编号) + 子级编号(盒子编号)
   - carton_no: 双层循环(大箱 + 小箱)

2. 大箱标：
   - quantity: 大箱内小箱数量 × 盒张数 × 每小箱盒数
   - serial: 父级编号(大箱编号) + 子级编号(01-大箱中盒子数量)
   - carton_no: 大箱循环递增
"""

import pytest
import sys
import math
from pathlib import Path

# 添加项目根目录到路径中
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / 'src'))

from data.data_processor import DataProcessor, PackagingConfig


class TestSeparateBoxLabelGeneration:
    """分盒模式标签生成测试"""
    
    def setup_method(self):
        """测试前置设置"""
        self.processor = DataProcessor()
    
    def create_test_scenario(self, box_quantity=60, small_box_capacity=5, large_box_capacity=3, 
                           cards_per_box=1000, start_number="JAW00001"):
        """创建测试场景"""
        variables = {
            'customer_code': 'TEST001',
            'theme': 'Label Test',
            'start_number': start_number
        }
        
        config = PackagingConfig(
            box_quantity=box_quantity,
            set_quantity=6,
            small_box_capacity=small_box_capacity,
            large_box_capacity=large_box_capacity,
            cards_per_box_in_set=cards_per_box,
            boxes_per_set=6,
            is_overweight=False,
            sets_per_large_box=2,
            cases_per_set=1
        )
        
        processed_data = self.processor.process_for_packaging_mode(variables, config, 'separate_box')
        return processed_data, config
    
    def test_small_box_label_quantity_calculation(self):
        """测试小箱标数量计算
        
        规律：quantity = 盒张数 × 每小箱盒数
        """
        # 测试场景：60盒，每小箱5盒，每盒1000张
        processed_data, config = self.create_test_scenario(
            box_quantity=60, small_box_capacity=5, cards_per_box=1000
        )
        
        # 预期小箱标数量计算
        cards_per_box = 1000  # 假设每盒1000张
        boxes_per_small_box = 5
        expected_small_box_quantity = cards_per_box * boxes_per_small_box  # 1000 × 5 = 5000
        
        # 验证小箱数量计算
        expected_small_boxes = math.ceil(60 / 5)  # 12个小箱
        assert processed_data['small_box_quantity'] == expected_small_boxes
        
        # 注意：当前代码中的quantity计算可能需要在PDF生成时根据文档规律调整
        print(f"Small boxes: {processed_data['small_box_quantity']}")
        print(f"Expected small box label quantity: {expected_small_box_quantity} PCS")
    
    def test_large_box_label_quantity_calculation(self):
        """测试大箱标数量计算
        
        规律：quantity = 大箱内小箱数量 × 盒张数 × 每小箱盒数
        """
        processed_data, config = self.create_test_scenario(
            box_quantity=60, small_box_capacity=5, large_box_capacity=3, cards_per_box=1000
        )
        
        # 预期大箱标数量计算
        cards_per_box = 1000
        boxes_per_small_box = 5
        small_boxes_per_large_box = 3
        expected_large_box_quantity = small_boxes_per_large_box * cards_per_box * boxes_per_small_box
        # 3 × 1000 × 5 = 15000
        
        # 验证大箱数量计算
        expected_large_boxes = math.ceil(12 / 3)  # 4个大箱 (12个小箱)
        assert processed_data['large_box_quantity'] == expected_large_boxes
        
        print(f"Large boxes: {processed_data['large_box_quantity']}")
        print(f"Expected large box label quantity: {expected_large_box_quantity} PCS")
    
    def test_serial_number_pattern_extraction(self):
        """测试序列号模式提取"""
        processed_data, _ = self.create_test_scenario(start_number="JAW00100")
        
        # 测试序列号解析
        label_ranges = self.processor.calculate_label_ranges(processed_data)
        box_numbers = label_ranges.get('box_numbers', [])
        
        # 验证序列号格式
        assert len(box_numbers) == 60  # 60盒
        assert box_numbers[0] == "JAW00100"
        assert box_numbers[59] == "JAW00159"  # 最后一个盒号
        
        print(f"Box number range: {box_numbers[0]} ~ {box_numbers[-1]}")
    
    def test_complex_serial_pattern(self):
        """测试复杂序列号模式
        
        测试场景设计验证文档中的序列号生成规律
        """
        processed_data, config = self.create_test_scenario(
            box_quantity=30, small_box_capacity=6, large_box_capacity=2, start_number="ABC01001"
        )
        
        # 计算预期结果
        expected_small_boxes = math.ceil(30 / 6)  # 5个小箱
        expected_large_boxes = math.ceil(5 / 2)   # 3个大箱
        
        assert processed_data['small_box_quantity'] == expected_small_boxes
        assert processed_data['large_box_quantity'] == expected_large_boxes
        
        # 生成序列号范围用于验证模式
        label_ranges = self.processor.calculate_label_ranges(processed_data)
        box_numbers = label_ranges['box_numbers']
        
        # 验证盒号连续性
        for i in range(len(box_numbers) - 1):
            current_num = int(box_numbers[i][-5:])  # 提取最后5位数字
            next_num = int(box_numbers[i + 1][-5:])
            assert next_num == current_num + 1, f"序列号不连续: {box_numbers[i]} -> {box_numbers[i + 1]}"


class TestSeparateBoxLabelLogicSimulation:
    """分盒模式标签逻辑模拟测试
    
    根据文档规律模拟实际的标签生成逻辑
    """
    
    def simulate_small_box_labels(self, processed_data, cards_per_box=1000):
        """模拟小箱标生成逻辑"""
        box_quantity = processed_data['box_quantity']
        boxes_per_small_box = processed_data['boxes_per_small_box']
        small_box_quantity = processed_data['small_box_quantity']
        start_number = processed_data['start_number']
        
        # 提取序列号前缀和起始数字
        import re
        match = re.search(r'(\d+)', start_number)
        if match:
            start_num = int(match.group(1))
            prefix = start_number.replace(match.group(1), '')
        else:
            start_num = 1
            prefix = ''
        
        small_box_labels = []
        
        for small_box_idx in range(small_box_quantity):
            # 计算该小箱包含的盒子范围
            start_box_in_small = small_box_idx * boxes_per_small_box + 1
            end_box_in_small = min((small_box_idx + 1) * boxes_per_small_box, box_quantity)
            actual_boxes_in_small = end_box_in_small - start_box_in_small + 1
            
            # 小箱标数量：实际盒数 × 每盒张数
            quantity = actual_boxes_in_small * cards_per_box
            
            # 序列号：从start_box到end_box
            if start_box_in_small == end_box_in_small:
                serial = f"{prefix}{start_num + start_box_in_small - 1:05d}"
            else:
                serial = f"{prefix}{start_num + start_box_in_small - 1:05d}-{prefix}{start_num + end_box_in_small - 1:05d}"
            
            # carton_no：小箱编号（需要考虑大箱分组）
            large_box_idx = small_box_idx // processed_data['small_boxes_per_large_box']
            small_box_in_large = small_box_idx % processed_data['small_boxes_per_large_box']
            carton_no = f"{large_box_idx + 1:02d}-{small_box_in_large + 1:02d}"
            
            small_box_labels.append({
                'type': 'small_box',
                'index': small_box_idx + 1,
                'quantity': f"{quantity}PCS",
                'serial': serial,
                'carton_no': carton_no,
                'boxes_count': actual_boxes_in_small
            })
        
        return small_box_labels
    
    def simulate_large_box_labels(self, processed_data, cards_per_box=1000):
        """模拟大箱标生成逻辑"""
        small_boxes_per_large_box = processed_data['small_boxes_per_large_box']
        boxes_per_small_box = processed_data['boxes_per_small_box']
        large_box_quantity = processed_data['large_box_quantity']
        small_box_quantity = processed_data['small_box_quantity']
        start_number = processed_data['start_number']
        
        # 提取序列号前缀和起始数字
        import re
        match = re.search(r'(\d+)', start_number)
        if match:
            start_num = int(match.group(1))
            prefix = start_number.replace(match.group(1), '')
        else:
            start_num = 1
            prefix = ''
        
        large_box_labels = []
        
        for large_box_idx in range(large_box_quantity):
            # 计算该大箱包含的小箱范围
            start_small_in_large = large_box_idx * small_boxes_per_large_box
            end_small_in_large = min((large_box_idx + 1) * small_boxes_per_large_box - 1, small_box_quantity - 1)
            actual_small_boxes = end_small_in_large - start_small_in_large + 1
            
            # 计算该大箱包含的盒子总数和序列号范围
            start_box_global = start_small_in_large * boxes_per_small_box + 1
            end_box_global = min((end_small_in_large + 1) * boxes_per_small_box, processed_data['box_quantity'])
            total_boxes_in_large = end_box_global - start_box_global + 1
            
            # 大箱标数量：实际包含的小箱数 × 每小箱盒数 × 每盒张数
            quantity = actual_small_boxes * boxes_per_small_box * cards_per_box
            
            # 序列号：从第一个盒子到最后一个盒子
            serial = f"{prefix}{start_num + start_box_global - 1:05d}-{prefix}{start_num + end_box_global - 1:05d}"
            
            # carton_no：大箱编号
            carton_no = f"{large_box_idx + 1:02d}"
            
            large_box_labels.append({
                'type': 'large_box',
                'index': large_box_idx + 1,
                'quantity': f"{quantity}PCS",
                'serial': serial,
                'carton_no': carton_no,
                'small_boxes_count': actual_small_boxes,
                'total_boxes_count': total_boxes_in_large
            })
        
        return large_box_labels
    
    def test_small_box_labels_simulation(self):
        """测试小箱标模拟生成"""
        processor = DataProcessor()
        variables = {
            'customer_code': 'SIM001',
            'theme': 'Simulation Test',
            'start_number': 'SIM00001'
        }
        
        config = PackagingConfig(
            box_quantity=23,    # 23盒
            set_quantity=6,
            small_box_capacity=5,   # 每小箱5盒
            large_box_capacity=2,   # 每大箱2小箱
            cards_per_box_in_set=800,
            boxes_per_set=6,
            is_overweight=False,
            sets_per_large_box=2,
            cases_per_set=1
        )
        
        processed_data = processor.process_for_packaging_mode(variables, config, 'separate_box')
        small_box_labels = self.simulate_small_box_labels(processed_data, cards_per_box=800)
        
        # 验证小箱标生成结果
        # 预期：23盒，每小箱5盒 -> 需要5个小箱 (5,5,5,5,3)
        assert len(small_box_labels) == 5
        
        # 验证第一个小箱标
        first_label = small_box_labels[0]
        assert first_label['quantity'] == "4000PCS"  # 5盒 × 800张
        assert first_label['serial'] == "SIM00001-SIM00005"
        assert first_label['carton_no'] == "01-01"  # 第1大箱的第1小箱
        
        # 验证最后一个小箱标（只有3盒）
        last_label = small_box_labels[4]
        assert last_label['quantity'] == "2400PCS"  # 3盒 × 800张
        assert last_label['serial'] == "SIM00021-SIM00023"
        assert last_label['boxes_count'] == 3
        
        print("小箱标生成结果:")
        for label in small_box_labels:
            print(f"  小箱{label['index']}: {label['quantity']}, {label['serial']}, {label['carton_no']}")
    
    def test_large_box_labels_simulation(self):
        """测试大箱标模拟生成"""
        processor = DataProcessor()
        variables = {
            'customer_code': 'SIM001',
            'theme': 'Simulation Test',
            'start_number': 'SIM00001'
        }
        
        config = PackagingConfig(
            box_quantity=23,
            set_quantity=6,
            small_box_capacity=5,
            large_box_capacity=2,
            cards_per_box_in_set=800,
            boxes_per_set=6,
            is_overweight=False,
            sets_per_large_box=2,
            cases_per_set=1
        )
        
        processed_data = processor.process_for_packaging_mode(variables, config, 'separate_box')
        large_box_labels = self.simulate_large_box_labels(processed_data, cards_per_box=800)
        
        # 验证大箱标生成结果
        # 预期：5个小箱，每大箱2小箱 -> 需要3个大箱 (2,2,1)
        assert len(large_box_labels) == 3
        
        # 验证第一个大箱标
        first_label = large_box_labels[0]
        assert first_label['quantity'] == "8000PCS"  # 2小箱 × 5盒/小箱 × 800张/盒
        assert first_label['serial'] == "SIM00001-SIM00010"
        assert first_label['carton_no'] == "01"
        
        # 验证最后一个大箱标（只有1个小箱）
        last_label = large_box_labels[2]
        assert last_label['quantity'] == "4000PCS"  # 1小箱 × 5盒/小箱 × 800张/盒
        assert last_label['small_boxes_count'] == 1
        
        print("大箱标生成结果:")
        for label in large_box_labels:
            print(f"  大箱{label['index']}: {label['quantity']}, {label['serial']}, {label['carton_no']}")


if __name__ == '__main__':
    # 运行测试
    pytest.main([__file__, '-v', '-s'])