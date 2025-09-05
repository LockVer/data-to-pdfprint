#!/usr/bin/env python3
"""
批量更新所有模版的字体系统，使用统一的中文字体工具
"""

import os
import re
from pathlib import Path

def update_template_file(file_path):
    """更新单个模版文件"""
    print(f"正在更新: {file_path}")
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 1. 添加字体工具导入（如果还没有的话）
        if 'from .font_utils import get_chinese_font' not in content:
            # 找到import math后面，添加字体导入
            if 'import math' in content:
                content = content.replace(
                    'import math',
                    '''import math

# 导入统一的字体工具
try:
    from .font_utils import get_chinese_font
except ImportError:
    def get_chinese_font():
        return 'Helvetica' '''
                )
        
        # 2. 更新__init__方法，使用get_chinese_font()
        if 'self.chinese_font = self._register_chinese_font()' in content:
            content = content.replace(
                'self.chinese_font = self._register_chinese_font()',
                'self.chinese_font = get_chinese_font()'
            )
        
        # 3. 删除旧的字体注册方法
        old_font_method_pattern = r'def _register_chinese_font\(self\):.*?return \'Helvetica.*?\'\s*'
        content = re.sub(old_font_method_pattern, '', content, flags=re.MULTILINE | re.DOTALL)
        
        # 写回文件
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
            
        print(f"  ✅ 成功更新: {file_path}")
        return True
        
    except Exception as e:
        print(f"  ❌ 更新失败: {file_path} - {e}")
        return False

def main():
    """主函数"""
    template_dir = Path("src/template")
    
    # 需要更新的模版文件
    template_files = [
        "inner_case_template.py",
        "outer_case_template.py",
        "set_box_label_template.py",
        "set_box_inner_case_template.py", 
        "set_box_outer_case_template.py",
        "division_inner_case_template.py",
        "division_outer_case_template.py"
    ]
    
    print("🔧 开始批量更新中文字体系统...")
    
    success_count = 0
    total_count = len(template_files)
    
    for filename in template_files:
        file_path = template_dir / filename
        if file_path.exists():
            if update_template_file(file_path):
                success_count += 1
        else:
            print(f"  ⚠️ 文件不存在: {file_path}")
    
    print(f"\\n✅ 批量更新完成: {success_count}/{total_count} 个文件成功更新")

if __name__ == "__main__":
    main()