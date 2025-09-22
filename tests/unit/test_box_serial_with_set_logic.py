#!/usr/bin/env python3
"""
盒标Serial生成逻辑测试用例
测试新的父级编号为套、子级编号为盒的逻辑
"""

import sys
import os
import pytest

# 添加src目录到Python路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', 'src'))

from pdf.split_box.data_processor import SplitBoxDataProcessor


class TestBoxSerialWithSetLogic:
    """盒标Serial生成逻辑测试类"""
    
    def setup_method(self):
        """测试前准备"""
        self.processor = SplitBoxDataProcessor()
    
    # ========== 基础功能测试 ==========
    
    def test_basic_serial_generation(self):
        """测试基础Serial生成功能"""
        base_number = "DSK01001-01"
        boxes_per_set = 6
        
        # 测试第1套的第1个盒
        result = self.processor.generate_box_serial_with_set_logic(base_number, 1, boxes_per_set)
        assert result == "DSK01001-01"
        
        # 测试第1套的最后一个盒
        result = self.processor.generate_box_serial_with_set_logic(base_number, 6, boxes_per_set)
        assert result == "DSK01001-06"
        
        # 测试第2套的第1个盒
        result = self.processor.generate_box_serial_with_set_logic(base_number, 7, boxes_per_set)
        assert result == "DSK01002-01"
    
    def test_different_boxes_per_set(self):
        """测试不同盒/套参数"""
        base_number = "DSK01001-01"
        
        # 测试盒/套=4的情况
        test_cases_4_boxes = [
            (1, "DSK01001-01"),   # 第1套第1盒
            (4, "DSK01001-04"),   # 第1套第4盒
            (5, "DSK01002-01"),   # 第2套第1盒
            (8, "DSK01002-04"),   # 第2套第4盒
            (9, "DSK01003-01"),   # 第3套第1盒
        ]
        
        for box_num, expected in test_cases_4_boxes:
            result = self.processor.generate_box_serial_with_set_logic(base_number, box_num, 4)
            assert result == expected, f"盒#{box_num} (盒/套=4): 期望{expected}, 实际{result}"
    
    # ========== 边界条件测试 ==========
    
    def test_boundary_conditions(self):
        """测试边界条件"""
        base_number = "DSK01001-01"
        boxes_per_set = 5
        
        # 边界测试用例
        boundary_cases = [
            (1, "DSK01001-01"),    # 第一个盒（绝对边界）
            (5, "DSK01001-05"),    # 第1套最后一个盒
            (6, "DSK01002-01"),    # 跨套边界（第2套第1盒）
            (10, "DSK01002-05"),   # 第2套最后一个盒
            (11, "DSK01003-01"),   # 跨套边界（第3套第1盒）
        ]
        
        for box_num, expected in boundary_cases:
            result = self.processor.generate_box_serial_with_set_logic(base_number, box_num, boxes_per_set)
            assert result == expected, f"边界测试盒#{box_num}: 期望{expected}, 实际{result}"
    
    # ========== 特殊场景测试 ==========
    
    def test_one_box_per_set(self):
        """测试一套一盒的特殊情况"""
        base_number = "DSK01001-01"
        boxes_per_set = 1
        
        test_cases = [
            (1, "DSK01001-01"),   # 第1套第1盒
            (2, "DSK01002-01"),   # 第2套第1盒
            (3, "DSK01003-01"),   # 第3套第1盒
            (10, "DSK01010-01"),  # 第10套第1盒
        ]
        
        for box_num, expected in test_cases:
            result = self.processor.generate_box_serial_with_set_logic(base_number, box_num, boxes_per_set)
            assert result == expected, f"一套一盒测试盒#{box_num}: 期望{expected}, 实际{result}"
    
    def test_large_numbers(self):
        """测试大数值情况"""
        base_number = "DSK01001-01"
        boxes_per_set = 100
        
        # 测试大数值场景
        large_cases = [
            (100, "DSK01001-100"),  # 第1套第100盒
            (101, "DSK01002-01"),   # 第2套第1盒
            (250, "DSK01003-50"),   # 第3套第50盒
            (1000, "DSK01010-100"), # 第10套第100盒
        ]
        
        for box_num, expected in large_cases:
            result = self.processor.generate_box_serial_with_set_logic(base_number, box_num, boxes_per_set)
            assert result == expected, f"大数值测试盒#{box_num}: 期望{expected}, 实际{result}"
    
    def test_different_base_number_formats(self):
        """测试不同的基准序列号格式"""
        boxes_per_set = 3
        
        # 不同格式的base_number测试
        format_cases = [
            ("ABC12345-01", 1, "ABC12345-01"),
            ("ABC12345-01", 3, "ABC12345-03"),
            ("ABC12345-01", 4, "ABC12346-01"),
            ("XYZ99999-01", 6, "XYZ100000-03"),  # 盒6：第2套第3盒，主号=99999+1=100000
            ("XYZ99999-01", 7, "XYZ100001-01"),  # 盒7：第3套第1盒，主号=99999+2=100001
        ]
        
        for base_number, box_num, expected in format_cases:
            result = self.processor.generate_box_serial_with_set_logic(base_number, box_num, boxes_per_set)
            assert result == expected, f"格式测试 {base_number} 盒#{box_num}: 期望{expected}, 实际{result}"
    
    # ========== 数学逻辑验证测试 ==========
    
    def test_divisible_scenarios(self):
        """测试整除场景（总盒数是盒/套的整倍数）"""
        base_number = "DSK01001-01"
        boxes_per_set = 4
        total_boxes = 12  # 恰好3套
        
        # 验证每套的完整性
        expected_results = [
            # 第1套
            (1, "DSK01001-01"), (2, "DSK01001-02"), (3, "DSK01001-03"), (4, "DSK01001-04"),
            # 第2套  
            (5, "DSK01002-01"), (6, "DSK01002-02"), (7, "DSK01002-03"), (8, "DSK01002-04"),
            # 第3套
            (9, "DSK01003-01"), (10, "DSK01003-02"), (11, "DSK01003-03"), (12, "DSK01003-04"),
        ]
        
        for box_num, expected in expected_results:
            result = self.processor.generate_box_serial_with_set_logic(base_number, box_num, boxes_per_set)
            assert result == expected, f"整除场景盒#{box_num}: 期望{expected}, 实际{result}"
    
    def test_non_divisible_scenarios(self):
        """测试非整除场景（最后一套盒数不满）"""
        base_number = "DSK01001-01"
        boxes_per_set = 5
        total_boxes = 13  # 2套满+1套不满(3个盒)
        
        expected_results = [
            # 第1套（满）
            (1, "DSK01001-01"), (2, "DSK01001-02"), (3, "DSK01001-03"), 
            (4, "DSK01001-04"), (5, "DSK01001-05"),
            # 第2套（满）
            (6, "DSK01002-01"), (7, "DSK01002-02"), (8, "DSK01002-03"),
            (9, "DSK01002-04"), (10, "DSK01002-05"),
            # 第3套（不满，只有3个盒）
            (11, "DSK01003-01"), (12, "DSK01003-02"), (13, "DSK01003-03"),
        ]
        
        for box_num, expected in expected_results:
            result = self.processor.generate_box_serial_with_set_logic(base_number, box_num, boxes_per_set)
            assert result == expected, f"非整除场景盒#{box_num}: 期望{expected}, 实际{result}"
    
    # ========== 逻辑一致性验证 ==========
    
    def test_logic_consistency(self):
        """测试逻辑一致性：同一套内的盒子主号相同，不同套主号递增"""
        base_number = "DSK01001-01"
        boxes_per_set = 4
        
        # 连续生成多个Serial，验证逻辑一致性
        results = []
        for box_num in range(1, 13):  # 测试3套共12个盒
            serial = self.processor.generate_box_serial_with_set_logic(base_number, box_num, boxes_per_set)
            results.append((box_num, serial))
        
        # 验证同套内主号相同
        assert results[0][1].split('-')[0] == results[3][1].split('-')[0]  # 第1套：盒1和盒4
        assert results[4][1].split('-')[0] == results[7][1].split('-')[0]  # 第2套：盒5和盒8
        assert results[8][1].split('-')[0] == results[11][1].split('-')[0] # 第3套：盒9和盒12
        
        # 验证不同套主号递增
        main1 = int(results[0][1].split('-')[0][-5:])   # 第1套主号
        main2 = int(results[4][1].split('-')[0][-5:])   # 第2套主号  
        main3 = int(results[8][1].split('-')[0][-5:])   # 第3套主号
        
        assert main2 == main1 + 1, f"第2套主号应该比第1套+1: {main1} → {main2}"
        assert main3 == main2 + 1, f"第3套主号应该比第2套+1: {main2} → {main3}"
    
    # ========== 性能测试 ==========
    
    def test_performance(self):
        """简单的性能测试：确保大量调用不会有明显性能问题"""
        import time
        
        base_number = "DSK01001-01"
        boxes_per_set = 100
        
        start_time = time.time()
        
        # 生成10000个Serial
        for box_num in range(1, 10001):
            self.processor.generate_box_serial_with_set_logic(base_number, box_num, boxes_per_set)
        
        elapsed = time.time() - start_time
        
        # 应该在1秒内完成（很宽松的性能要求）
        assert elapsed < 1.0, f"性能测试失败：10000次调用耗时{elapsed:.3f}秒"


# ========== 辅助函数 ==========

def run_visual_test():
    """可视化测试：打印一些测试结果供人工验证"""
    processor = SplitBoxDataProcessor()
    base_number = "DSK01001-01"
    
    print("\n" + "="*60)
    print("盒标Serial生成逻辑可视化测试")
    print("="*60)
    
    # 测试场景1：每套6个盒，生成前15个盒的Serial
    print(f"\n📋 场景1: 每套6个盒，生成前15个盒")
    boxes_per_set = 6
    for box_num in range(1, 16):
        serial = processor.generate_box_serial_with_set_logic(base_number, box_num, boxes_per_set)
        set_num = ((box_num - 1) // boxes_per_set) + 1
        box_in_set = ((box_num - 1) % boxes_per_set) + 1
        print(f"  盒#{box_num:2d}: {serial} (第{set_num}套第{box_in_set}盒)")
    
    # 测试场景2：一套一盒，生成前10个盒
    print(f"\n📋 场景2: 每套1个盒，生成前10个盒")
    boxes_per_set = 1
    for box_num in range(1, 11):
        serial = processor.generate_box_serial_with_set_logic(base_number, box_num, boxes_per_set)
        print(f"  盒#{box_num:2d}: {serial}")
    
    print("\n" + "="*60)


if __name__ == "__main__":
    # 运行可视化测试
    run_visual_test()
    
    # 运行pytest测试
    print("\n🧪 运行pytest测试...")
    pytest.main([__file__, "-v"])