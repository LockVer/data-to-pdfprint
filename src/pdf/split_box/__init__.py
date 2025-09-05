"""
Split Box Template Package
分盒模板包
"""

import os
import importlib.util

# 使用importlib导入模板避免相对导入问题
template_path = os.path.join(os.path.dirname(__file__), "template.py")
spec = importlib.util.spec_from_file_location("split_box_template", template_path)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
SplitBoxTemplate = module.SplitBoxTemplate

__all__ = ['SplitBoxTemplate']