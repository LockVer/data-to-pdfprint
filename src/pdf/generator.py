"""
PDF生成器

使用ReportLab生成PDF文档
"""

from reportlab.lib.pagesizes import A4, LETTER, A5
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm, cm
from reportlab.lib.colors import black, blue, red
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pathlib import Path
from typing import List, Dict, Any, Union
import os
import platform


class PDFGenerator:
    """
    PDF生成器类
    """

    def __init__(self):
        """
        初始化PDF生成器
        """
        self.page_size = A4
        self.margin = 20 * mm
        self.font_name = "Helvetica"
        self.font_size = 12
        self.chinese_font_name = "SimHei"
        self._register_chinese_font()

    def set_page_size(self, size: str):
        """
        设置页面大小

        Args:
            size: 页面大小 ('A4', 'LETTER', 'A5')
        """
        size_map = {
            'A4': A4,
            'LETTER': LETTER,
            'A5': A5
        }
        
        if size.upper() in size_map:
            self.page_size = size_map[size.upper()]
        else:
            raise ValueError(f"不支持的页面大小: {size}")

    def generate_from_template(self, template, data: Dict[str, Any], output_path: str):
        """
        根据模板和数据生成PDF

        Args:
            template: 模板对象
            data: 数据字典
            output_path: 输出文件路径
        """
        if not output_path:
            raise ValueError("输出路径不能为空")
        
        # 确保输出目录存在
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        
        # 创建PDF画布
        c = canvas.Canvas(str(output_file), pagesize=self.page_size)
        
        try:
            # 获取页面尺寸
            width, height = self.page_size
            
            # 设置基本字体
            c.setFont(self.font_name, self.font_size)
            
            # 如果有模板，使用模板渲染
            if template and hasattr(template, 'render'):
                template.render(c, data, width, height)
            else:
                # 使用默认布局
                self._render_default_layout(c, data, width, height)
            
            # 保存PDF
            c.save()
            
        except Exception as e:
            raise Exception(f"生成PDF失败: {e}")

    def batch_generate(self, template, data_list: List[Dict[str, Any]], output_dir: str) -> List[str]:
        """
        批量生成PDF

        Args:
            template: 模板对象
            data_list: 数据列表
            output_dir: 输出目录

        Returns:
            生成的文件路径列表
        """
        if not data_list:
            return []
        
        output_files = []
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        for i, data in enumerate(data_list):
            # 生成文件名
            filename = f"label_{i+1:03d}.pdf"
            if 'name' in data:
                filename = f"label_{data['name']}_{i+1:03d}.pdf"
            
            file_path = output_path / filename
            
            try:
                self.generate_from_template(template, data, str(file_path))
                output_files.append(str(file_path))
            except Exception as e:
                print(f"生成第{i+1}个PDF失败: {e}")
                continue
        
        return output_files

    def create_label_pdf(self, data: Dict[str, Any], output_path: str):
        """
        创建标签PDF

        Args:
            data: 标签数据
            output_path: 输出路径
        """
        self.generate_from_template(None, data, output_path)

    def _render_default_layout(self, canvas_obj, data: Dict[str, Any], width: float, height: float):
        """
        渲染默认布局

        Args:
            canvas_obj: 画布对象
            data: 数据
            width: 页面宽度
            height: 页面高度
        """
        # 计算起始位置
        x_start = self.margin
        y_start = height - self.margin - self.font_size
        
        line_height = self.font_size + 5
        y_pos = y_start
        
        # 渲染标题
        title = "数据标签"
        if self._has_chinese(title):
            canvas_obj.setFont(self.chinese_font_name, self.font_size + 4)
        else:
            canvas_obj.setFont(self.font_name, self.font_size + 4)
        canvas_obj.drawString(x_start, y_pos, title)
        y_pos -= line_height * 2
        
        # 渲染数据字段
        canvas_obj.setFont(self.font_name, self.font_size)
        for key, value in data.items():
            text = f"{key}: {value}"
            # 如果包含中文，使用中文字体
            if self._has_chinese(text):
                canvas_obj.setFont(self.chinese_font_name, self.font_size)
            else:
                canvas_obj.setFont("Helvetica", self.font_size)
            canvas_obj.drawString(x_start, y_pos, text)
            y_pos -= line_height

    def _register_chinese_font(self):
        """
        注册中文字体
        """
        try:
            # 根据操作系统选择字体路径
            system = platform.system()
            font_paths = []
            
            if system == "Darwin":  # macOS
                font_paths = [
                    "/System/Library/Fonts/Supplemental/Songti.ttc",
                    "/System/Library/Fonts/STHeiti Light.ttc",
                    "/System/Library/Fonts/STHeiti Medium.ttc",
                    "/Library/Fonts/SimHei.ttf"
                ]
            elif system == "Windows":
                font_paths = [
                    "C:/Windows/Fonts/simhei.ttf",
                    "C:/Windows/Fonts/simsun.ttc",
                    "C:/Windows/Fonts/msyh.ttc"
                ]
            elif system == "Linux":
                font_paths = [
                    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
                    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
                ]
            
            # 尝试注册第一个可用的字体
            for font_path in font_paths:
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(TTFont(self.chinese_font_name, font_path))
                        self.font_name = self.chinese_font_name
                        return
                    except Exception as e:
                        continue
            
            # 如果没有找到中文字体，保持默认字体
            print("警告: 未找到中文字体，使用默认字体")
            
        except Exception as e:
            print(f"注册中文字体失败: {e}")

    def _has_chinese(self, text: str) -> bool:
        """
        检查文本是否包含中文字符
        """
        for char in str(text):
            if '\u4e00' <= char <= '\u9fff':
                return True
        return False
