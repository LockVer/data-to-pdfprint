#!/usr/bin/env python3
"""
Serial逻辑快速测试
专注于核心功能的快速验证
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.pdf.split_box.data_processor import SplitBoxDataProcessor


def test_serial_quick():
    """快速测试核心Serial功能"""
    processor = SplitBoxDataProcessor()
    
    print("🚀 Serial逻辑快速测试")
    print("=" * 50)
    
    tests = [
        # 测试1：多套分一箱（你遇到的问题）
        {
            "name": "多套分一箱",
            "func": processor.generate_set_based_large_box_serial_range,
            "args": (1, "JAW01001-01", 3, 8, 1, 12),
            "expected": "JAW01001-01-JAW01003-02"
        },
        
        # 测试2：一套分多箱（之前修复的问题）
        {
            "name": "一套分多箱", 
            "func": processor.generate_set_based_large_box_serial_range,
            "args": (2, "DSK01001-01", 15, 8, 1, 30),
            "expected": "DSK01001-09-DSK01001-15"
        },
        
        # 测试3：一套分一箱（格式问题）
        {
            "name": "一套分一箱",
            "func": processor.generate_set_based_large_box_serial_range,
            "args": (1, "BAT01001-01", 6, 6, 1, 12),
            "expected": "BAT01001-01-BAT01001-06"
        },
        
        # 测试4：小箱标多套分一箱
        {
            "name": "小箱多套分一箱",
            "func": processor.generate_set_based_small_box_serial_range,
            "args": (1, "JAW01001-01", 4, 8, 16),
            "expected": "JAW01001-01-JAW01002-04"
        },
        
        # 测试5：小箱标一套分多箱
        {
            "name": "小箱一套分多箱",
            "func": processor.generate_set_based_small_box_serial_range,
            "args": (2, "DSK01001-01", 9, 3, 18),
            "expected": "DSK01001-04-DSK01001-06"
        },
        
        # 测试6：单盒Serial
        {
            "name": "单盒Serial",
            "func": processor.generate_set_based_box_serial,
            "args": (25, "DSK01001-01", 10),
            "expected": "DSK01003-05"
        }
    ]
    
    passed = 0
    failed = 0
    
    for test in tests:
        try:
            print(f"\n🧪 {test['name']}")
            result = test['func'](*test['args'])
            
            if result == test['expected']:
                print(f"✅ 通过: {result}")
                passed += 1
            else:
                print(f"❌ 失败: 期望 {test['expected']}, 实际 {result}")
                failed += 1
                
        except Exception as e:
            print(f"❌ 异常: {e}")
            failed += 1
    
    print(f"\n🏁 测试完成: {passed} 通过, {failed} 失败")
    return failed == 0


if __name__ == "__main__":
    success = test_serial_quick()
    exit(0 if success else 1)