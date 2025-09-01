"""
PDF生成器测试
"""

import pytest
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
from src.pdf.generator import PDFGenerator


class TestPDFGenerator:
    """PDFGenerator类的测试用例"""
    
    def test_init(self):
        """测试初始化"""
        generator = PDFGenerator()
        assert generator is not None
    
    @patch('src.pdf.generator.canvas')
    def test_generate_from_template_success(self, mock_canvas):
        """测试成功生成PDF"""
        # 模拟画布对象
        mock_canvas_obj = MagicMock()
        mock_canvas.Canvas.return_value = mock_canvas_obj
        
        generator = PDFGenerator()
        template = Mock()
        template.render.return_value = "rendered content"
        
        data = {'name': '张三', 'age': 25}
        output_path = "output/test.pdf"
        
        generator.generate_from_template(template, data, output_path)
        
        # 验证Canvas被正确调用 - 包含pagesize参数
        mock_canvas.Canvas.assert_called_once_with("output/test.pdf", pagesize=generator.page_size)
        mock_canvas_obj.save.assert_called_once()
    
    def test_generate_from_template_invalid_path(self):
        """测试无效输出路径"""
        generator = PDFGenerator()
        template = Mock()
        data = {'name': '张三'}
        
        with pytest.raises(Exception):
            generator.generate_from_template(template, data, "")
    
    @patch('src.pdf.generator.canvas')
    def test_batch_generate_success(self, mock_canvas):
        """测试批量生成PDF"""
        mock_canvas_obj = MagicMock()
        mock_canvas.Canvas.return_value = mock_canvas_obj
        
        generator = PDFGenerator()
        template = Mock()
        template.render.return_value = "rendered content"
        
        data_list = [
            {'name': '张三', 'age': 25},
            {'name': '李四', 'age': 30}
        ]
        output_dir = "output/"
        
        generator.batch_generate(template, data_list, output_dir)
        
        # 验证为每个数据项生成了PDF
        assert mock_canvas.Canvas.call_count == 2
        assert mock_canvas_obj.save.call_count == 2
    
    def test_batch_generate_empty_data(self):
        """测试空数据列表的批量生成"""
        generator = PDFGenerator()
        template = Mock()
        
        result = generator.batch_generate(template, [], "output/")
        assert result == []
    
    def test_create_label_pdf(self):
        """测试创建标签PDF"""
        generator = PDFGenerator()
        
        label_data = {
            'title': '产品标签',
            'name': '张三',
            'code': 'P001',
            'date': '2024-01-01'
        }
        
        # 测试方法存在性
        assert hasattr(generator, 'create_label_pdf')
    
    def test_set_page_size(self):
        """测试设置页面大小"""
        from reportlab.lib.pagesizes import A4, LETTER, A5
        
        generator = PDFGenerator()
        
        # 测试不同页面大小
        size_map = {'A4': A4, 'LETTER': LETTER, 'A5': A5}
        
        for size_name, expected_size in size_map.items():
            generator.set_page_size(size_name)
            assert generator.page_size == expected_size