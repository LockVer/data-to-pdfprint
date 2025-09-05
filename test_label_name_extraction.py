#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试标签名称提取脚本
"""

import sys
import os
from pathlib import Path

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from data.excel_reader import ExcelReader
from template.division_inner_case_template import DivisionInnerCaseTemplate

def test_label_name_extraction():
    """测试标签名称提取功能"""
    print("=" * 80)
    print("开始测试标签名称提取功能...")
    print("=" * 80)
    
    # Excel文件路径
    excel_file = "/Users/heye/Desktop/大白鲨 xlsx.xlsx"
    
    if not Path(excel_file).exists():
        print(f"❌ Excel文件不存在: {excel_file}")
        return
    
    try:
        # 1. 读取Excel数据 - 使用单元格地址格式
        print(f"📖 读取Excel文件: {excel_file}")
        reader = ExcelReader(excel_file)
        excel_data = reader.read_data_by_cell_address()
        
        print(f"✅ 成功读取Excel数据，共 {len(excel_data)} 个单元格")
        
        # 2. 创建模板实例并测试提取
        template = DivisionInnerCaseTemplate()
        
        print("\n" + "="*50)
        print("开始测试标签名称提取:")
        print("="*50)
        
        # 直接调用提取方法
        extracted_data = template._search_label_name_data(excel_data)
        
        print("="*50)
        print(f"📊 最终提取结果: '{extracted_data}'")
        print("="*50)
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n" + "=" * 80)
    print("标签名称提取测试完成")
    print("=" * 80)

if __name__ == "__main__":
    test_label_name_extraction()