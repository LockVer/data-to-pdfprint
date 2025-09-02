#!/usr/bin/env python3
"""
测试修正后的盒标生成
"""

import sys
import os
from pathlib import Path

# 添加src目录到路径
sys.path.insert(0, 'src')

from src.pdf.generator import PDFGenerator

def test_box_labels():
    """测试盒标生成"""
    
    # 从实际Excel文件读取数据
    import pandas as pd
    df = pd.read_excel('/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx', header=None)
    
    # 提取数据
    test_data = {
        '客户编码': str(df.iloc[3,0]),
        '主题': str(df.iloc[3,1]),  # B4单元格的产品名称
        '排列要求': str(df.iloc[3,2]),
        '订单数量': str(df.iloc[3,3]),
        '总张数': str(df.iloc[3,5])
    }
    
    print("📊 从Excel提取的数据:")
    for key, value in test_data.items():
        print(f"   {key}: {value}")
    
    # 用户参数
    packaging_params = {
        '张/盒': 2850,    # 用户输入：每盒2850张
        '盒/小箱': 1,     # 用户输入：每小箱1盒
        '小箱/大箱': 2,   # 用户输入：每大箱2小箱
        '选择外观': '外观一'  # 测试外观一
    }
    
    print(f"\n📦 包装参数:")
    for key, value in packaging_params.items():
        print(f"   {key}: {value}")
    
    # 计算盒数
    total_pieces = int(test_data['总张数'])
    pieces_per_box = packaging_params['张/盒']
    total_boxes = -(-total_pieces // pieces_per_box)  # 向上取整
    
    print(f"\n🔢 计算结果:")
    print(f"   总张数: {total_pieces}")
    print(f"   张/盒: {pieces_per_box}")
    print(f"   总盒数: {total_boxes} (应该生成{total_boxes}页PDF)")
    
    # 输出目录
    output_dir = "./test_box_output"
    
    # 创建PDF生成器
    generator = PDFGenerator()
    
    try:
        print(f"\n🔄 开始生成盒标PDF (外观一)...")
        generated_files = generator.create_multi_level_pdfs(
            test_data, 
            packaging_params, 
            output_dir
        )
        
        print("\n✅ 生成完成!")
        for label_type, file_path in generated_files.items():
            if label_type == "盒标":
                print(f"  📄 {label_type}: {Path(file_path).name}")
                size = os.path.getsize(file_path)
                print(f"      文件大小: {size} 字节")
                print(f"      应包含: {total_boxes}页")
        
        return True
        
    except Exception as e:
        print(f"❌ 生成失败: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    test_box_labels()