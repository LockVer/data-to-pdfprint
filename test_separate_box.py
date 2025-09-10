#!/usr/bin/env python3
"""
分盒模式数据生成逻辑测试运行器

使用方法：
    python test_separate_box.py              # 运行所有测试
    python test_separate_box.py --basic      # 只运行基础逻辑测试
    python test_separate_box.py --labels     # 只运行标签生成测试
    python test_separate_box.py --demo       # 运行演示示例
"""

import sys
import argparse
from pathlib import Path

# 添加项目源码路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root / 'src'))

from data.data_processor import DataProcessor, PackagingConfig


def run_basic_demo():
    """运行基础演示"""
    print("\n" + "="*60)
    print("分盒模式数据生成逻辑演示")
    print("="*60)
    
    processor = DataProcessor()
    
    # 演示场景设置
    variables = {
        'customer_code': 'DEMO001',
        'theme': '演示测试主题',
        'start_number': 'JAW00001'
    }
    
    scenarios = [
        {
            'name': '基础场景',
            'box_quantity': 100,
            'small_box_capacity': 6,
            'large_box_capacity': 4,
            'description': '100盒卡片，每小箱装6盒，每大箱装4小箱'
        },
        {
            'name': '边界场景',
            'box_quantity': 37,
            'small_box_capacity': 8,
            'large_box_capacity': 3,
            'description': '37盒卡片，每小箱装8盒，每大箱装3小箱'
        },
        {
            'name': '大数量场景',
            'box_quantity': 5000,
            'small_box_capacity': 50,
            'large_box_capacity': 20,
            'description': '5000盒卡片，每小箱装50盒，每大箱装20小箱'
        }
    ]
    
    for i, scenario in enumerate(scenarios, 1):
        print(f"\n场景 {i}: {scenario['name']}")
        print(f"描述: {scenario['description']}")
        print("-" * 40)
        
        config = PackagingConfig(
            box_quantity=scenario['box_quantity'],
            set_quantity=6,
            small_box_capacity=scenario['small_box_capacity'],
            large_box_capacity=scenario['large_box_capacity'],
            cards_per_box_in_set=630,
            boxes_per_set=6,
            is_overweight=False,
            sets_per_large_box=2,
            cases_per_set=1
        )
        
        # 处理数据
        result = processor.process_for_packaging_mode(variables, config, 'separate_box')
        
        # 显示结果
        print(f"输入参数:")
        print(f"  盒数量: {scenario['box_quantity']}")
        print(f"  每小箱盒数: {scenario['small_box_capacity']}")
        print(f"  每大箱小箱数: {scenario['large_box_capacity']}")
        
        print(f"\n计算结果:")
        print(f"  小箱数量: {result['small_box_quantity']}")
        print(f"  大箱数量: {result['large_box_quantity']}")
        
        print(f"\n标签规格:")
        specs = result['label_specifications']
        print(f"  盒标数量: {specs['box_labels']}")
        print(f"  小箱标数量: {specs['small_box_labels']}")
        print(f"  大箱标数量: {specs['large_box_labels']}")
        
        # 计算包装效率
        total_capacity = result['large_box_quantity'] * scenario['large_box_capacity'] * scenario['small_box_capacity']
        efficiency = (scenario['box_quantity'] / total_capacity) * 100 if total_capacity > 0 else 0
        print(f"\n包装效率: {efficiency:.1f}% ({scenario['box_quantity']}/{total_capacity})")


def run_label_simulation_demo():
    """运行标签模拟演示"""
    print("\n" + "="*60)
    print("分盒模式标签生成模拟演示")
    print("="*60)
    
    # 导入标签模拟功能
    sys.path.insert(0, str(Path(__file__).parent / 'tests'))
    from test_separate_box_labels import TestSeparateBoxLabelLogicSimulation
    
    processor = DataProcessor()
    sim_test = TestSeparateBoxLabelLogicSimulation()
    
    # 设置演示场景
    variables = {
        'customer_code': 'LABEL001',
        'theme': '标签演示',
        'start_number': 'LAB00100'
    }
    
    config = PackagingConfig(
        box_quantity=27,      # 27盒
        set_quantity=6,
        small_box_capacity=6, # 每小箱6盒
        large_box_capacity=2, # 每大箱2小箱
        cards_per_box_in_set=1000,
        boxes_per_set=6,
        is_overweight=False,
        sets_per_large_box=2,
        cases_per_set=1
    )
    
    print(f"演示场景: 27盒卡片，每盒1000张，每小箱6盒，每大箱2小箱")
    print("-" * 60)
    
    # 处理数据
    processed_data = processor.process_for_packaging_mode(variables, config, 'separate_box')
    
    print(f"计算结果:")
    print(f"  需要小箱: {processed_data['small_box_quantity']}个")
    print(f"  需要大箱: {processed_data['large_box_quantity']}个")
    
    # 生成小箱标
    small_box_labels = sim_test.simulate_small_box_labels(processed_data, cards_per_box=1000)
    
    print(f"\n小箱标生成结果 (共{len(small_box_labels)}个):")
    print("-" * 60)
    for label in small_box_labels:
        print(f"  小箱{label['index']:2d}: {label['quantity']:>8s} | {label['serial']:>15s} | 箱号:{label['carton_no']} | 含{label['boxes_count']}盒")
    
    # 生成大箱标
    large_box_labels = sim_test.simulate_large_box_labels(processed_data, cards_per_box=1000)
    
    print(f"\n大箱标生成结果 (共{len(large_box_labels)}个):")
    print("-" * 60)
    for label in large_box_labels:
        print(f"  大箱{label['index']:2d}: {label['quantity']:>8s} | {label['serial']:>15s} | 箱号:{label['carton_no']} | 含{label['small_boxes_count']}小箱")


def run_tests():
    """运行完整测试套件"""
    print("\n" + "="*60)
    print("运行分盒模式完整测试套件")
    print("="*60)
    
    try:
        import pytest
        
        # 运行测试
        test_files = [
            'tests/test_separate_box_mode.py',
            'tests/test_separate_box_labels.py'
        ]
        
        args = ['-v', '-s', '--tb=short']
        args.extend(test_files)
        
        return pytest.main(args)
        
    except ImportError:
        print("错误: 未安装 pytest")
        print("请运行: pip install pytest")
        return 1


def main():
    parser = argparse.ArgumentParser(description='分盒模式测试运行器')
    parser.add_argument('--basic', action='store_true', help='只运行基础逻辑测试')
    parser.add_argument('--labels', action='store_true', help='只运行标签生成测试')
    parser.add_argument('--demo', action='store_true', help='运行演示示例')
    parser.add_argument('--all', action='store_true', help='运行完整测试套件')
    
    args = parser.parse_args()
    
    if args.demo:
        run_basic_demo()
        run_label_simulation_demo()
    elif args.basic:
        try:
            import pytest
            return pytest.main(['tests/test_separate_box_mode.py', '-v', '-s'])
        except ImportError:
            print("错误: 未安装 pytest。请运行: pip install pytest")
            return 1
    elif args.labels:
        try:
            import pytest
            return pytest.main(['tests/test_separate_box_labels.py', '-v', '-s'])
        except ImportError:
            print("错误: 未安装 pytest。请运行: pip install pytest")
            return 1
    elif args.all:
        return run_tests()
    else:
        # 默认运行演示
        print("分盒模式测试工具")
        print("使用 --help 查看所有选项")
        print("\n运行演示示例...")
        run_basic_demo()


if __name__ == '__main__':
    sys.exit(main())