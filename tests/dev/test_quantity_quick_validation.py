#!/usr/bin/env python3
"""
Quantity计算逻辑快速验证脚本

这是一个轻量级的验证脚本，用于快速检查quantity计算逻辑是否正常工作
适合在开发过程中进行快速验证，不需要运行完整的测试套件
"""

import sys
import os

# 添加项目路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.pdf.split_box.data_processor import SplitBoxDataProcessor


def test_basic_functionality():
    """测试基本功能"""
    print("🔍 测试基本功能...")
    
    processor = SplitBoxDataProcessor()
    
    # 测试1: 小箱quantity计算
    result = processor.calculate_actual_quantity_for_small_box(1, 730, 2, 10)
    expected = 730 * 2  # 1460
    assert result == expected, f"小箱量计算错误: 期望{expected}, 实际{result}"
    print(f"✅ 小箱quantity计算: {result} PCS")
    
    # 测试2: 大箱quantity计算
    result = processor.calculate_actual_quantity_for_large_box(1, 730, 2, 4, 20)
    expected = 730 * 8  # 5840 (2*4=8盒/大箱)
    assert result == expected, f"大箱量计算错误: 期望{expected}, 实际{result}"
    print(f"✅ 大箱quantity计算: {result} PCS")


def test_boundary_cases():
    """测试边界情况"""
    print("\n🔍 测试边界情况...")
    
    processor = SplitBoxDataProcessor()
    
    # 测试1: 最后一个容器包含较少盒子
    result = processor.calculate_actual_quantity_for_small_box(5, 100, 3, 13)
    expected = 100 * 1  # 最后一个小箱只包含1盒（盒13）
    assert result == expected, f"边界情况计算错误: 期望{expected}, 实际{result}"
    print(f"✅ 边界情况处理: {result} PCS (最后1盒)")
    
    # 测试2: 容器编号超出范围
    result = processor.calculate_actual_quantity_for_small_box(10, 100, 3, 13)
    expected = 0  # 超出范围应该返回0
    assert result == expected, f"超范围处理错误: 期望{expected}, 实际{result}"
    print(f"✅ 超范围处理: {result} PCS")


def test_real_world_scenario():
    """测试真实世界场景"""
    print("\n🔍 测试真实场景...")
    
    processor = SplitBoxDataProcessor()
    
    # 用户实际案例：109500张，730张/盒，150盒总数
    pieces_per_box = 730
    total_boxes = 150  # ceil(109500 / 730)
    
    # 二级模式：8盒/大箱
    boxes_per_large_box = 8
    
    # 测试第1个大箱
    result = processor.calculate_actual_quantity_for_large_box(1, pieces_per_box, boxes_per_large_box, 1, total_boxes)
    expected = 730 * 8  # 5840
    assert result == expected, f"真实场景第1个大箱错误: 期望{expected}, 实际{result}"
    print(f"✅ 第1个大箱: {result} PCS")
    
    # 测试最后一个大箱（第19个，包含6盒：145-150）
    result = processor.calculate_actual_quantity_for_large_box(19, pieces_per_box, boxes_per_large_box, 1, total_boxes)
    expected = 730 * 6  # 4380
    assert result == expected, f"真实场景最后大箱错误: 期望{expected}, 实际{result}"
    print(f"✅ 最后大箱: {result} PCS (实际6盒)")


def test_consistency():
    """测试一致性"""
    print("\n🔍 测试计算一致性...")
    
    processor = SplitBoxDataProcessor()
    
    # 三级模式：验证大箱quantity等于其包含的小箱quantity之和
    pieces_per_box = 500
    boxes_per_small_box = 3
    small_boxes_per_large_box = 4
    total_boxes = 24  # 刚好2个大箱
    
    # 第1个大箱应该包含前4个小箱
    large_box_quantity = processor.calculate_actual_quantity_for_large_box(
        1, pieces_per_box, boxes_per_small_box, small_boxes_per_large_box, total_boxes
    )
    
    # 前4个小箱的总量
    small_box_total = 0
    for i in range(1, 5):
        quantity = processor.calculate_actual_quantity_for_small_box(
            i, pieces_per_box, boxes_per_small_box, total_boxes
        )
        small_box_total += quantity
    
    assert large_box_quantity == small_box_total, \
        f"大箱与小箱总量不一致: 大箱{large_box_quantity}, 小箱总计{small_box_total}"
    print(f"✅ 一致性验证: 大箱{large_box_quantity} = 小箱总计{small_box_total}")


def test_performance_sanity():
    """测试性能合理性"""
    print("\n🔍 测试性能合理性...")
    
    processor = SplitBoxDataProcessor()
    
    import time
    
    # 测试1000次计算的时间
    start_time = time.time()
    for i in range(1000):
        result = processor.calculate_actual_quantity_for_small_box(i + 1, 700, 5, 10000)
    elapsed = time.time() - start_time
    
    rate = 1000 / elapsed
    print(f"✅ 性能测试: 1000次计算耗时{elapsed:.3f}秒，速率{rate:.0f}次/秒")
    
    # 合理的性能期望：应该能达到至少100次/秒
    assert rate > 100, f"性能过低: {rate:.0f}次/秒"


def main():
    """主函数"""
    print("🚀 Quantity计算逻辑快速验证")
    print("=" * 50)
    
    try:
        test_basic_functionality()
        test_boundary_cases()
        test_real_world_scenario()
        test_consistency()
        test_performance_sanity()
        
        print("\n" + "=" * 50)
        print("🎉 所有快速验证测试通过！")
        print("✅ Quantity计算逻辑工作正常")
        return True
        
    except AssertionError as e:
        print(f"\n❌ 验证失败: {e}")
        return False
    except Exception as e:
        print(f"\n💥 验证过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)