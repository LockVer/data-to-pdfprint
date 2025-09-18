#!/usr/bin/env python3
"""
快速测试脚本 - 验证核心计算逻辑
运行: python quick_test.py
"""

import math
import sys
import os
from datetime import datetime

# 添加项目路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.pdf.split_box.data_processor import SplitBoxDataProcessor

def test_your_scenario():
    """测试你提到的具体场景"""
    print("🧪 测试你的具体场景")
    print("-" * 40)
    
    processor = SplitBoxDataProcessor()
    
    # 你的参数
    params = {
        "张/盒": 730,
        "盒/套": 15, 
        "盒/小箱": 8,  # 实际是盒/大箱
        "小箱/大箱": 1,
        "是否有小箱": False
    }
    
    # 手动计算期望值
    total_pieces = 109500
    pieces_per_box = 730
    boxes_per_set = 15
    boxes_per_large_box = 8
    
    total_boxes = math.ceil(total_pieces / pieces_per_box)
    total_sets = math.ceil(total_boxes / boxes_per_set)
    large_boxes_per_set = math.ceil(boxes_per_set / boxes_per_large_box)
    total_large_boxes = total_sets * large_boxes_per_set
    
    print(f"📊 计算过程:")
    print(f"   总张数: {total_pieces}")
    print(f"   张/盒: {pieces_per_box}")
    print(f"   总盒数: {total_boxes} = ceil({total_pieces} ÷ {pieces_per_box})")
    print(f"   盒/套: {boxes_per_set}")
    print(f"   总套数: {total_sets} = ceil({total_boxes} ÷ {boxes_per_set})")
    print(f"   盒/大箱: {boxes_per_large_box}")
    print(f"   每套大箱数: {large_boxes_per_set} = ceil({boxes_per_set} ÷ {boxes_per_large_box})")
    print(f"   总大箱数: {total_large_boxes} = {total_sets} × {large_boxes_per_set}")
    
    print(f"\n📦 生成Carton No (前10个):")
    carton_nos = []
    for i in range(1, min(total_large_boxes + 1, 11)):
        carton_no = processor.calculate_carton_range_for_large_box(i, boxes_per_set, boxes_per_large_box, total_sets)
        carton_nos.append(carton_no)
        print(f"   大箱 #{i}: {carton_no}")
    
    print(f"\n📦 最后几个Carton No:")
    if total_large_boxes > 10:
        for i in range(max(total_large_boxes - 2, 11), total_large_boxes + 1):
            carton_no = processor.calculate_carton_range_for_large_box(i, boxes_per_set, boxes_per_large_box, total_sets)
            print(f"   大箱 #{i}: {carton_no}")
    
    # 验证结果
    expected_sequence = ["1-1", "1-2", "2-1", "2-2", "3-1", "3-2", "4-1", "4-2", "5-1", "5-2"]
    expected_last = ["10-1", "10-2"]
    
    print(f"\n✅ 验证结果:")
    print(f"   期望总大箱数: 20, 实际: {total_large_boxes}")
    print(f"   期望前10个: {expected_sequence}")
    print(f"   实际前10个: {carton_nos}")
    print(f"   期望最后2个: {expected_last}")
    
    # 验证结果
    verification_passed = total_large_boxes == 20 and carton_nos == expected_sequence
    if verification_passed:
        print(f"🎉 测试通过! 逻辑正确!")
        verification_result = "✅ 通过"
    else:
        print(f"❌ 测试失败! 请检查逻辑!")
        verification_result = "❌ 失败"
    
    # 生成完整的Carton No序列
    all_carton_nos = []
    for i in range(1, total_large_boxes + 1):
        carton_no = processor.calculate_carton_range_for_large_box(i, boxes_per_set, boxes_per_large_box, total_sets)
        all_carton_nos.append(carton_no)
    
    # 返回测试结果
    return {
        "name": "你的具体场景测试 (二级模式)",
        "params": params,
        "total_pieces": total_pieces,
        "total_boxes": total_boxes,
        "total_sets": total_sets,
        "large_boxes_per_set": large_boxes_per_set,
        "total_large_boxes": total_large_boxes,
        "large_carton_nos": all_carton_nos,
        "verification": verification_result
    }

def test_three_level_mode():
    """测试三级模式"""
    print(f"\n\n🧪 测试三级模式 (有小箱)")
    print("-" * 40)
    
    processor = SplitBoxDataProcessor()
    
    params = {
        "张/盒": 1000,
        "盒/套": 6,
        "盒/小箱": 2,
        "小箱/大箱": 2,
        "是否有小箱": True
    }
    
    # 计算
    total_pieces = 12000
    pieces_per_box = 1000
    boxes_per_set = 6
    boxes_per_small_box = 2
    small_boxes_per_large_box = 2
    
    total_boxes = math.ceil(total_pieces / pieces_per_box)
    total_sets = math.ceil(total_boxes / boxes_per_set)
    small_boxes_per_set = math.ceil(boxes_per_set / boxes_per_small_box)
    large_boxes_per_set = math.ceil(small_boxes_per_set / small_boxes_per_large_box)
    total_small_boxes = total_sets * small_boxes_per_set
    total_large_boxes = total_sets * large_boxes_per_set
    
    print(f"📊 三级模式计算:")
    print(f"   总盒数: {total_boxes}, 总套数: {total_sets}")
    print(f"   每套小箱数: {small_boxes_per_set}, 总小箱数: {total_small_boxes}")
    print(f"   每套大箱数: {large_boxes_per_set}, 总大箱数: {total_large_boxes}")
    
    print(f"\n📦 小箱Carton No:")
    for i in range(1, total_small_boxes + 1):
        carton_no = processor.calculate_carton_number_for_small_box(i, boxes_per_set, boxes_per_small_box)
        print(f"   小箱 #{i}: {carton_no}")
    
    print(f"\n📦 大箱Carton No:")
    boxes_per_large_box = boxes_per_small_box * small_boxes_per_large_box
    
    # 生成所有Carton No
    small_carton_nos = []
    large_carton_nos = []
    
    for i in range(1, total_small_boxes + 1):
        carton_no = processor.calculate_carton_number_for_small_box(i, boxes_per_set, boxes_per_small_box)
        small_carton_nos.append(carton_no)
    
    for i in range(1, total_large_boxes + 1):
        carton_no = processor.calculate_carton_range_for_large_box(i, boxes_per_set, boxes_per_large_box, total_sets)
        large_carton_nos.append(carton_no)
        print(f"   大箱 #{i}: {carton_no}")
    
    # 返回测试结果
    return {
        "name": "三级模式测试 (有小箱)",
        "params": params,
        "total_pieces": total_pieces,
        "total_boxes": total_boxes,
        "total_sets": total_sets,
        "small_boxes_per_set": small_boxes_per_set,
        "large_boxes_per_set": large_boxes_per_set,
        "total_small_boxes": total_small_boxes,
        "total_large_boxes": total_large_boxes,
        "small_carton_nos": small_carton_nos,
        "large_carton_nos": large_carton_nos,
        "verification": "✅ 三级模式测试完成"
    }

def export_quick_test_results(test_results):
    """导出快速测试结果"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"quick_test_results_{timestamp}.md"
    
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(f"# 快速测试结果报告\n\n")
            f.write(f"**测试时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            
            for i, result in enumerate(test_results, 1):
                f.write(f"## {i}. {result['name']}\n\n")
                
                # 测试参数
                f.write(f"**测试参数**:\n")
                for key, value in result['params'].items():
                    f.write(f"- {key}: {value}\n")
                f.write(f"- 总张数: {result['total_pieces']}\n\n")
                
                # 计算过程
                f.write(f"**计算过程**:\n")
                f.write(f"- 总盒数: {result['total_boxes']} = ceil({result['total_pieces']} ÷ {result['params']['张/盒']})\n")
                f.write(f"- 总套数: {result['total_sets']} = ceil({result['total_boxes']} ÷ {result['params']['盒/套']})\n")
                if 'small_boxes_per_set' in result:
                    f.write(f"- 每套小箱数: {result['small_boxes_per_set']}\n")
                    f.write(f"- 每套大箱数: {result['large_boxes_per_set']}\n")
                    f.write(f"- 总小箱数: {result['total_small_boxes']}\n")
                else:
                    f.write(f"- 每套大箱数: {result['large_boxes_per_set']}\n")
                f.write(f"- 总大箱数: {result['total_large_boxes']}\n\n")
                
                # Carton No结果
                if 'small_carton_nos' in result:
                    f.write(f"**小箱Carton No (前10个)**:\n")
                    for j, carton in enumerate(result['small_carton_nos'][:10], 1):
                        f.write(f"- 小箱 #{j}: {carton}\n")
                    f.write(f"\n")
                
                f.write(f"**大箱Carton No (前10个)**:\n")
                for j, carton in enumerate(result['large_carton_nos'][:10], 1):
                    f.write(f"- 大箱 #{j}: {carton}\n")
                
                if len(result['large_carton_nos']) > 10:
                    f.write(f"\n**大箱Carton No (最后几个)**:\n")
                    for j in range(max(len(result['large_carton_nos']) - 2, 10), len(result['large_carton_nos'])):
                        f.write(f"- 大箱 #{j+1}: {result['large_carton_nos'][j]}\n")
                
                # 验证结果
                if 'verification' in result:
                    f.write(f"\n**验证结果**: {result['verification']}\n")
                
                f.write(f"\n---\n\n")
        
        print(f"\n📄 快速测试结果已导出: {filename}")
        
    except Exception as e:
        print(f"\n❌ 结果导出失败: {str(e)}")

def main():
    """主函数"""
    print("🚀 快速测试 Carton Number 逻辑")
    print("=" * 50)
    
    test_results = []
    
    # 测试你的具体场景
    result1 = test_your_scenario()
    if result1:
        test_results.append(result1)
    
    # 测试三级模式
    result2 = test_three_level_mode()
    if result2:
        test_results.append(result2)
    
    print(f"\n🏁 快速测试完成!")
    
    # 导出测试结果
    if test_results:
        export_quick_test_results(test_results)

if __name__ == "__main__":
    # 临时禁用详细调试输出，保持测试清晰
    import sys
    from io import StringIO
    
    old_stdout = sys.stdout
    sys.stdout = StringIO()
    
    try:
        # 先恢复输出用于测试
        sys.stdout = old_stdout
        main()
    except Exception as e:
        sys.stdout = old_stdout
        print(f"测试异常: {str(e)}")