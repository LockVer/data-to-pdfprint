#!/usr/bin/env python3
"""
测试Serial范围格式显示
验证一套入一箱场景下Serial始终显示为范围形式
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.pdf.split_box.data_processor import SplitBoxDataProcessor


def test_serial_range_format():
    """测试Serial范围格式"""
    processor = SplitBoxDataProcessor()
    
    print("🧪 测试Serial范围格式")
    print("=" * 60)
    
    base_number = "BAT-01001-01"
    
    # 测试1：一套入一箱（起始和结束Serial相同）
    print("\n📋 测试：一套入一箱（起始=结束）")
    print("参数：盒/套=6, 盒/大箱=6")
    print("期望：BAT-01001-01-BAT-01001-06")
    
    boxes_per_set = 6
    boxes_per_small_box = 6
    small_boxes_per_large_box = 1
    total_boxes = 12
    
    result = processor.generate_set_based_large_box_serial_range(
        1, base_number, boxes_per_set, 
        boxes_per_small_box, small_boxes_per_large_box, total_boxes
    )
    print(f"实际结果: {result}")
    
    # 测试2：跨套显示（起始和结束Serial不同）
    print("\n📋 测试：多套分一箱（起始≠结束）") 
    print("参数：盒/套=3, 盒/大箱=8")
    print("期望：BAT-01001-01-BAT-01003-02")
    
    boxes_per_set = 3
    boxes_per_small_box = 8
    total_boxes = 24
    
    result = processor.generate_set_based_large_box_serial_range(
        1, base_number, boxes_per_set, 
        boxes_per_small_box, small_boxes_per_large_box, total_boxes
    )
    print(f"实际结果: {result}")
    
    # 测试3：小箱标也测试一下
    print("\n📋 测试：小箱标范围格式")
    print("参数：盒/套=4, 盒/小箱=4")
    print("期望：BAT-01001-01-BAT-01001-04")
    
    boxes_per_set = 4
    boxes_per_small_box = 4
    total_boxes = 8
    
    result = processor.generate_set_based_small_box_serial_range(
        1, base_number, boxes_per_set, boxes_per_small_box, total_boxes
    )
    print(f"实际结果: {result}")
    
    print("\n🏁 测试完成")


if __name__ == "__main__":
    test_serial_range_format()