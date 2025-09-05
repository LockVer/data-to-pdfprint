#!/usr/bin/env python3
"""
修复模版文件中的语法问题
"""

import os
import re
from pathlib import Path

def fix_template_file(file_path):
    """修复单个模版文件的语法"""
    print(f"正在修复: {file_path}")
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 修复孤立的 except 块
        # 查找并删除孤立的 except Exception as e: 块
        orphaned_except_pattern = r'\n\s*except Exception as e:\s*print\(f"字体注册失败.*?\n'
        content = re.sub(orphaned_except_pattern, '\n', content, flags=re.MULTILINE | re.DOTALL)
        
        # 确保 __init__ 方法有正确的内容
        if 'def __init__(self):' in content and 'pass' in content:
            # 替换简单的 pass 为完整的初始化
            init_pattern = r'def __init__\(self\):\s*"""初始化模板"""\s*[^}]*pass'
            replacement = '''def __init__(self):
        """初始化模板"""
        self.chinese_font = get_chinese_font()
        
        # 颜色定义 (CMYK)
        self.colors = {
            'black': CMYKColor(0, 0, 0, 100),
            'gray': CMYKColor(0, 0, 0, 60),
            'light_gray': CMYKColor(0, 0, 0, 20)
        }'''
            content = re.sub(init_pattern, replacement, content, flags=re.MULTILINE | re.DOTALL)
        
        # 写回文件
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
            
        print(f"  ✅ 成功修复: {file_path}")
        return True
        
    except Exception as e:
        print(f"  ❌ 修复失败: {file_path} - {e}")
        return False

def main():
    """主函数"""
    template_dir = Path("src/template")
    
    # 需要修复的模版文件
    template_files = [
        "inner_case_template.py",
        "outer_case_template.py", 
        "set_box_inner_case_template.py", 
        "set_box_outer_case_template.py",
        "division_outer_case_template.py"
    ]
    
    print("🔧 开始修复语法错误...")
    
    success_count = 0
    total_count = len(template_files)
    
    for filename in template_files:
        file_path = template_dir / filename
        if file_path.exists():
            if fix_template_file(file_path):
                success_count += 1
        else:
            print(f"  ⚠️ 文件不存在: {file_path}")
    
    print(f"\\n✅ 语法修复完成: {success_count}/{total_count} 个文件成功修复")

if __name__ == "__main__":
    main()