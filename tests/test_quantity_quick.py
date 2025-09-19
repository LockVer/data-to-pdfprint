#!/usr/bin/env python3
"""
Quantity逻辑快速测试
专注于核心功能的快速验证，与serial/carton测试风格保持一致
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.pdf.split_box.data_processor import SplitBoxDataProcessor


def test_quantity_quick():
    """快速测试核心Quantity功能"""
    processor = SplitBoxDataProcessor()
    
    print("🚀 Quantity逻辑快速测试")
    print("=" * 50)
    
    tests = [
        # 测试1: 小箱quantity计算
        {
            "name": "小箱基础计算",
            "func": processor.calculate_actual_quantity_for_small_box,
            "args": (1, 730, 2, 150),  # 第1个小箱，730张/盒，2盒/小箱，总150盒
            "expected": 1460,  # 730 × 2 = 1460
            "desc": "小箱#1: 2盒 × 730张/盒 = 1460张"
        },
        
        # 测试2: 大箱quantity计算（无套模式）
        {
            "name": "大箱均匀分配",
            "func": processor.calculate_actual_quantity_for_large_box,
            "args": (1, 730, 2, 4, 150, None),  # 第1个大箱，无套模式
            "expected": 5840,  # 730 × 8 = 5840
            "desc": "大箱#1 (无套): 8盒 × 730张/盒 = 5840张"
        },
        
        # 测试3: 大箱quantity计算（套盒模式：一套分多箱）
        {
            "name": "大箱套盒模式",
            "func": processor.calculate_actual_quantity_for_large_box,
            "args": (2, 730, 2, 4, 150, 15),  # 第2个大箱，15盒/套
            "expected": 5110,  # 730 × 7 = 5110（完成第一套的后7盒）
            "desc": "大箱#2 (套盒): 7盒 × 730张/盒 = 5110张"
        },
        
        # 测试4: 边界情况 - 最后容器包含不完整盒数
        {
            "name": "最后容器边界",
            "func": processor.calculate_actual_quantity_for_small_box,
            "args": (5, 100, 3, 13),  # 最后小箱，总13盒
            "expected": 100,  # 只有1盒，100 × 1 = 100
            "desc": "最后小箱: 1盒 × 100张/盒 = 100张"
        },
        
        # 测试5: 边界情况 - 超出范围
        {
            "name": "超出范围处理",
            "func": processor.calculate_actual_quantity_for_small_box,
            "args": (10, 100, 3, 13),  # 第10个小箱，但总共只有13盒
            "expected": 0,  # 超出范围，返回0
            "desc": "超出范围: 0张"
        },
        
        # 测试6: 验证与serial分配的一致性（用户场景）
        {
            "name": "用户场景验证",
            "func": processor.calculate_actual_quantity_for_large_box,
            "args": (1, 730, 2, 4, 150, 15),  # 用户实际参数
            "expected": 5840,  # 730 × 8 = 5840
            "desc": "用户场景大箱#1: 8盒 × 730张/盒 = 5840张"
        },
        
        # 测试7: 验证套盒模式第二个大箱
        {
            "name": "用户场景验证2",
            "func": processor.calculate_actual_quantity_for_large_box,
            "args": (2, 730, 2, 4, 150, 15),  # 用户实际参数
            "expected": 5110,  # 730 × 7 = 5110
            "desc": "用户场景大箱#2: 7盒 × 730张/盒 = 5110张"
        }
    ]
    
    passed = 0
    total = len(tests)
    
    for i, test in enumerate(tests, 1):
        print(f"\n🔍 测试 {i}/{total}: {test['name']}")
        try:
            result = test["func"](*test["args"])
            expected = test["expected"]
            
            if result == expected:
                print(f"✅ {test['desc']} - 通过")
                passed += 1
            else:
                print(f"❌ {test['desc']} - 失败")
                print(f"   期望: {expected}, 实际: {result}")
                
        except Exception as e:
            print(f"❌ {test['desc']} - 异常: {e}")
    
    print(f"\n{'='*50}")
    print(f"📊 测试结果: {passed}/{total} 通过")
    
    if passed == total:
        print("🎉 所有测试通过！")
        return True
    else:
        print("⚠️ 部分测试失败")
        return False


def test_quantity_consistency():
    """测试quantity计算与其他逻辑的一致性"""
    processor = SplitBoxDataProcessor()
    
    print("\n🔄 Quantity一致性验证")
    print("=" * 50)
    
    # 测试大箱与小箱量的一致性
    pieces_per_box = 500
    boxes_per_small_box = 3
    small_boxes_per_large_box = 4
    total_boxes = 12
    
    # 计算一个大箱的quantity
    large_box_quantity = processor.calculate_actual_quantity_for_large_box(
        1, pieces_per_box, boxes_per_small_box, small_boxes_per_large_box, total_boxes, None
    )
    
    # 计算这个大箱对应的所有小箱quantity总和
    small_box_quantities = []
    for small_box_num in range(1, small_boxes_per_large_box + 1):
        qty = processor.calculate_actual_quantity_for_small_box(
            small_box_num, pieces_per_box, boxes_per_small_box, total_boxes
        )
        small_box_quantities.append(qty)
    
    small_box_total = sum(small_box_quantities)
    
    print(f"大箱#1 quantity: {large_box_quantity}")
    print(f"小箱quantity总和: {small_box_total} (小箱: {small_box_quantities})")
    
    if large_box_quantity == small_box_total:
        print("✅ 大箱与小箱quantity计算一致")
        return True
    else:
        print("❌ 大箱与小箱quantity计算不一致")
        return False


def test_quantity_performance():
    """简单的性能验证"""
    processor = SplitBoxDataProcessor()
    
    print("\n⚡ 性能基准测试")
    print("=" * 50)
    
    import time
    
    # 测试1000次quantity计算的性能
    start_time = time.time()
    
    for i in range(1000):
        processor.calculate_actual_quantity_for_large_box(
            i % 20 + 1, 730, 2, 4, 150, 15
        )
    
    end_time = time.time()
    duration = end_time - start_time
    rate = 1000 / duration
    
    print(f"1000次大箱quantity计算耗时: {duration:.3f}秒")
    print(f"计算速率: {rate:.0f}次/秒")
    
    if rate > 1000:  # 期望 > 1000次/秒
        print("✅ 性能测试通过")
        return True
    else:
        print("⚠️ 性能可能需要优化")
        return False


if __name__ == "__main__":
    """运行所有快速测试"""
    print("🧪 开始Quantity逻辑快速测试套件")
    print("="*60)
    
    results = []
    
    # 运行核心功能测试
    results.append(test_quantity_quick())
    
    # 运行一致性测试
    results.append(test_quantity_consistency())
    
    # 运行性能测试
    results.append(test_quantity_performance())
    
    # 总结
    passed_suites = sum(results)
    total_suites = len(results)
    
    print(f"\n{'='*60}")
    print(f"🏁 测试套件完成: {passed_suites}/{total_suites} 套件通过")
    
    if passed_suites == total_suites:
        print("🎉 所有测试套件通过！Quantity逻辑运行正常")
        exit(0)
    else:
        print("⚠️ 部分测试套件失败，请检查quantity计算逻辑")
        exit(1)