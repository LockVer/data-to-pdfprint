"""
PDF生成器

使用ReportLab生成PDF文档
"""

from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch, cm
from reportlab.lib.colors import black, blue, red
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from typing import Dict, Any, List
import os
from pathlib import Path


class PDFGenerator:
    """
    PDF生成器类
    """
    
    def __init__(self):
        """
        初始化PDF生成器
        """
        self.styles = getSampleStyleSheet()
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=16,
            spaceAfter=20,
            alignment=1  # Center alignment
        )
        self.content_style = ParagraphStyle(
            'CustomContent',
            parent=self.styles['Normal'],
            fontSize=12,
            spaceAfter=10
        )
    
    def generate_box_labels(self, data: Dict[str, Any], output_path: str, template: str = 'regular'):
        """
        生成盒标PDF
        
        Args:
            data: 处理后的数据
            output_path: 输出文件路径
            template: 模板类型
        """
        doc = SimpleDocTemplate(output_path, pagesize=A4)
        story = []
        
        # 标题
        title = f"盒标 - {data.get('theme', 'Unknown')}"
        story.append(Paragraph(title, self.title_style))
        story.append(Spacer(1, 20))
        
        # 基本信息表
        basic_info = [
            ['客户编码', data.get('customer_code', '')],
            ['主题', data.get('theme', '')],
            ['开始号', data.get('start_number', '')],
            ['包装模式', self._get_mode_display_name(data.get('packaging_mode', ''))],
            ['盒数量', str(data.get('box_quantity', 0))]
        ]
        
        if template == 'set_box':
            basic_info.extend([
                ['套数量', str(data.get('set_quantity', 6))],
                ['总套数', str(data.get('total_sets', 0))]
            ])
        
        basic_table = Table(basic_info, colWidths=[3*cm, 5*cm])
        basic_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        story.append(basic_table)
        story.append(Spacer(1, 30))
        
        # 生成具体的盒标内容
        story.extend(self._generate_box_label_content(data, template))
        
        doc.build(story)
    
    def generate_case_labels(self, data: Dict[str, Any], output_path: str, template: str = 'regular'):
        """
        生成箱标PDF
        
        Args:
            data: 处理后的数据
            output_path: 输出文件路径
            template: 模板类型
        """
        doc = SimpleDocTemplate(output_path, pagesize=A4)
        story = []
        
        # 标题
        title = f"箱标 - {data.get('theme', 'Unknown')}"
        story.append(Paragraph(title, self.title_style))
        story.append(Spacer(1, 20))
        
        # 箱装信息
        if template == 'regular':
            case_info = [
                ['包装方式', '一盒入箱再两盒入一箱'],
                ['小箱数量', str(data.get('small_box_quantity', 0))],
                ['大箱数量', str(data.get('large_box_quantity', 0))],
                ['每小箱盒数', str(data.get('boxes_per_small_box', 1))],
                ['每大箱小箱数', str(data.get('small_boxes_per_large_box', 2))]
            ]
        elif template == 'separate_box':
            case_info = [
                ['包装方式', '一盒入一小箱 再两小箱入一大箱'],
                ['小箱数量', str(data.get('small_box_quantity', 0))],
                ['大箱数量', str(data.get('large_box_quantity', 0))],
                ['每小箱盒数', str(data.get('boxes_per_small_box', 1))],
                ['每大箱小箱数', str(data.get('small_boxes_per_large_box', 2))]
            ]
        elif template == 'set_box':
            case_info = [
                ['包装方式', '六盒为一套 六盒入一小箱 再两套入一大箱'],
                ['套数', str(data.get('total_sets', 0))],
                ['小箱数量', str(data.get('small_box_quantity', 0))],
                ['大箱数量', str(data.get('large_box_quantity', 0))],
                ['每套盒数', str(data.get('boxes_per_set', 6))],
                ['每小箱套数', str(data.get('sets_per_small_box', 1))],
                ['每大箱套数', str(data.get('sets_per_large_box', 2))]
            ]
        
        case_table = Table(case_info, colWidths=[4*cm, 4*cm])
        case_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.lightcyan),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        story.append(case_table)
        story.append(Spacer(1, 30))
        
        # 生成具体的箱标内容
        story.extend(self._generate_case_label_content(data, template))
        
        doc.build(story)
    
    def _generate_box_label_content(self, data: Dict[str, Any], template: str) -> List:
        """生成盒标具体内容"""
        from ..data.data_processor import DataProcessor
        
        processor = DataProcessor()
        label_ranges = processor.calculate_label_ranges(data)
        
        content = []
        content.append(Paragraph("盒标明细", self.title_style))
        
        # 显示盒号范围
        box_numbers = label_ranges.get('box_numbers', [])
        if box_numbers:
            if len(box_numbers) <= 20:  # 如果数量不多，显示全部
                numbers_display = ', '.join(box_numbers)
            else:  # 数量太多，显示范围
                numbers_display = f"{box_numbers[0]} ~ {box_numbers[-1]} (共{len(box_numbers)}个)"
            
            content.append(Paragraph(f"盒号: {numbers_display}", self.content_style))
        
        if template == 'set_box':
            set_numbers = label_ranges.get('set_numbers', [])
            if set_numbers:
                set_display = ', '.join(set_numbers)
                content.append(Paragraph(f"套号: {set_display}", self.content_style))
        
        return content
    
    def _generate_case_label_content(self, data: Dict[str, Any], template: str) -> List:
        """生成箱标具体内容"""
        content = []
        content.append(Paragraph("箱标明细", self.title_style))
        
        # 根据不同模式生成不同的箱标内容
        if template == 'regular':
            content.append(Paragraph("常规包装模式箱标", self.content_style))
            for i in range(data.get('large_box_quantity', 0)):
                content.append(Paragraph(f"大箱 {i+1}: 包含2个小箱", self.content_style))
        
        elif template == 'separate_box':
            content.append(Paragraph("分盒包装模式箱标", self.content_style))
            for i in range(data.get('large_box_quantity', 0)):
                content.append(Paragraph(f"大箱 {i+1}: 包含2个小箱，每小箱1盒", self.content_style))
        
        elif template == 'set_box':
            content.append(Paragraph("套盒包装模式箱标", self.content_style))
            for i in range(data.get('large_box_quantity', 0)):
                content.append(Paragraph(f"大箱 {i+1}: 包含2套，每套6盒", self.content_style))
        
        return content
    
    def _get_mode_display_name(self, mode: str) -> str:
        """获取模式的中文显示名称"""
        mode_names = {
            'regular': '常规',
            'separate_box': '分盒', 
            'set_box': '套盒'
        }
        return mode_names.get(mode, mode)
    
    def generate_from_template(self, template, data, output_path: str):
        """
        根据模板和数据生成PDF
        
        Args:
            template: 模板对象
            data: 数据
            output_path: 输出文件路径
        """
        # 为了向后兼容保留此方法
        self.generate_box_labels(data, output_path, template)
    
    def batch_generate(self, template, data_list, output_dir: str):
        """
        批量生成PDF
        
        Args:
            template: 模板对象
            data_list: 数据列表
            output_dir: 输出目录
        """
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        for i, data in enumerate(data_list):
            box_file = output_path / f"box_labels_{i+1}.pdf"
            case_file = output_path / f"case_labels_{i+1}.pdf"
            
            self.generate_box_labels(data, str(box_file), template)
            self.generate_case_labels(data, str(case_file), template)