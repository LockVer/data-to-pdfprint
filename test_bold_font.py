#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试粗体字体注册脚本
"""

import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from template.font_utils import get_chinese_font, get_chinese_bold_font

def test_bold_font_registration():
    """测试粗体字体注册"""
    print("=" * 50)
    print("开始测试粗体字体注册...")
    print("=" * 50)
    
    # 测试常规字体
    print("\n1. 测试常规字体:")
    regular_font = get_chinese_font()
    print(f"常规字体: {regular_font}")
    
    # 测试粗体字体
    print("\n2. 测试粗体字体:")
    bold_font = get_chinese_bold_font()
    print(f"粗体字体: {bold_font}")
    
    # 比较结果
    if bold_font != regular_font:
        print(f"\n✅ 成功: 粗体字体 ({bold_font}) 与常规字体 ({regular_font}) 不同")
    else:
        print(f"\n⚠️  注意: 粗体字体与常规字体相同，可能使用常规字体代替")
    
    print("\n" + "=" * 50)
    print("粗体字体测试完成")
    print("=" * 50)

if __name__ == "__main__":
    test_bold_font_registration()