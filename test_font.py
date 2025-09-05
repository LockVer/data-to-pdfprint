#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试字体注册脚本
"""

import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from template.font_utils import register_chinese_font, get_chinese_font

def test_font_registration():
    """测试字体注册"""
    print("=" * 50)
    print("开始测试字体注册...")
    print("=" * 50)
    
    # 重置字体缓存
    from template import font_utils
    font_utils._chinese_font_cache = None
    
    # 测试字体注册
    print("\n1. 测试字体注册:")
    font_name = register_chinese_font()
    print(f"注册的字体: {font_name}")
    
    # 测试缓存
    print("\n2. 测试字体缓存:")
    cached_font = get_chinese_font()
    print(f"缓存的字体: {cached_font}")
    
    # 检查是否成功注册微软雅黑
    if "MicrosoftYaHei" in font_name:
        print("\n✅ 成功: 微软雅黑字体已注册!")
    else:
        print(f"\n⚠️  警告: 未能注册微软雅黑，当前使用: {font_name}")
    
    print("\n" + "=" * 50)
    print("字体测试完成")
    print("=" * 50)

if __name__ == "__main__":
    test_font_registration()