"""
PDF生成器 - 重构版本
使用委托模式将不同模板的逻辑分离到独立文件中
"""

from typing import Dict, Any
import importlib.util
import sys
import os


def _import_template_class(template_dir: str, class_name: str):
    """使用importlib安全导入模板类"""
    template_path = os.path.join(os.path.dirname(__file__), template_dir, "template.py")
    
    spec = importlib.util.spec_from_file_location(f"{template_dir}_template", template_path)
    if spec is None:
        raise ImportError(f"无法加载模板: {template_path}")
    
    module = importlib.util.module_from_spec(spec)
    
    # 添加路径到sys.modules以支持模板内的相对导入
    sys.modules[f"{template_dir}_template"] = module
    
    spec.loader.exec_module(module)
    return getattr(module, class_name)


# 模板类将在需要时延迟导入
_template_classes = {}


def _get_template_class(template_name: str):
    """延迟导入模板类"""
    if template_name not in _template_classes:
        if template_name == "regular":
            _template_classes[template_name] = _import_template_class("regular", "RegularTemplate")
        elif template_name == "split_box":
            _template_classes[template_name] = _import_template_class("split_box", "SplitBoxTemplate")
        elif template_name == "nested_box":
            _template_classes[template_name] = _import_template_class("nested_box", "NestedBoxTemplate")
        else:
            raise ValueError(f"未知的模板类型: {template_name}")
    
    return _template_classes[template_name]


class PDFGenerator:
    """
    PDF生成器主类 - 重构版本
    通过委托模式调用不同的模板类
    """

    def __init__(self, max_pages_per_file: int = 100):
        """
        初始化PDF生成器
        
        Args:
            max_pages_per_file: 每个PDF文件的最大页数限制
        """
        self.max_pages_per_file = max_pages_per_file
        
        # 模板实例将在需要时延迟创建
        self._regular_template = None
        self._split_box_template = None
        self._nested_box_template = None
    
    @property
    def regular_template(self):
        """延迟创建常规模板实例"""
        if self._regular_template is None:
            RegularTemplate = _get_template_class("regular")
            self._regular_template = RegularTemplate(self.max_pages_per_file)
        return self._regular_template
    
    @property
    def split_box_template(self):
        """延迟创建分盒模板实例"""
        if self._split_box_template is None:
            SplitBoxTemplate = _get_template_class("split_box")
            self._split_box_template = SplitBoxTemplate(self.max_pages_per_file)
        return self._split_box_template
    
    @property
    def nested_box_template(self):
        """延迟创建套盒模板实例"""
        if self._nested_box_template is None:
            NestedBoxTemplate = _get_template_class("nested_box")
            self._nested_box_template = NestedBoxTemplate(self.max_pages_per_file)
        return self._nested_box_template

    def create_multi_level_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """
        创建常规模板的多级标签PDF
        """
        return self.regular_template.create_multi_level_pdfs(data, params, output_dir, excel_file_path)

    def create_split_box_multi_level_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """
        Create multi-level PDF labels for split box template
        """
        return self.split_box_template.create_multi_level_pdfs(data, params, output_dir, excel_file_path)

    def create_nested_box_multi_level_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """
        Create multi-level PDF labels for nested box template
        """
        return self.nested_box_template.create_multi_level_pdfs(data, params, output_dir, excel_file_path)

    # 保持向后兼容的一些通用方法
    def set_page_size(self, size: str):
        """设置页面尺寸"""
        # 为保持兼容性，同步设置所有模板的页面尺寸
        self.regular_template.set_page_size(size) if hasattr(self.regular_template, 'set_page_size') else None
        self.split_box_template.set_page_size(size) if hasattr(self.split_box_template, 'set_page_size') else None
        self.nested_box_template.set_page_size(size) if hasattr(self.nested_box_template, 'set_page_size') else None