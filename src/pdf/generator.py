"""
PDF生成器

使用ReportLab生成PDF文档
"""

from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfgen import canvas
from reportlab.lib.colors import black, blue, red, green
from reportlab.lib.units import mm, cm, inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping
from pathlib import Path
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
        self.chinese_font = self._register_chinese_font()
        
    def _register_chinese_font(self):
        """
        注册中文字体
        """
        try:
            system = platform.system()
            
            # 尝试不同操作系统的中文字体路径
            font_paths = []
            
            if system == "Darwin":  # macOS
                font_paths = [
                    "/System/Library/Fonts/PingFang.ttc",
                    "/System/Library/Fonts/Supplemental/Arial Unicode MS.ttf",
                    "/System/Library/Fonts/Helvetica.ttc",
                    "/System/Library/Fonts/Arial.ttf"
                ]
            elif system == "Windows":
                font_paths = [
                    "C:/Windows/Fonts/msyh.ttc",  # 微软雅黑
                    "C:/Windows/Fonts/simhei.ttf",  # 黑体
                    "C:/Windows/Fonts/simsun.ttc"   # 宋体
                ]
            else:  # Linux
                font_paths = [
                    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
                ]
            
            # 尝试注册字体
            for font_path in font_paths:
                try:
                    if os.path.exists(font_path):
                        print(f"尝试加载字体: {font_path}")
                        if font_path.endswith('.ttc'):
                            # TTC文件需要指定字体索引，尝试多个索引
                            for index in range(5):  # 尝试前5个字体
                                try:
                                    pdfmetrics.registerFont(TTFont('ChineseFont', font_path, subfontIndex=index))
                                    addMapping('ChineseFont', 0, 0, 'ChineseFont')
                                    print(f"成功加载字体: {font_path} (索引: {index})")
                                    return 'ChineseFont'
                                except:
                                    continue
                        else:
                            pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                            addMapping('ChineseFont', 0, 0, 'ChineseFont')
                            print(f"成功加载字体: {font_path}")
                            return 'ChineseFont'
                except Exception as e:
                    print(f"加载字体失败 {font_path}: {e}")
                    continue
            
            # 尝试使用ReportLab内置的中文字体支持
            try:
                from reportlab.lib.fonts import addMapping
                from reportlab.pdfbase.cidfonts import UnicodeCIDFont
                pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
                return 'STSong-Light'
            except:
                pass
            
            # 最后的备用方案
            print("警告: 无法加载中文字体，将使用英文字体，中文可能显示异常")
            return 'Helvetica'
            
        except Exception as e:
            print(f"注册中文字体失败: {e}")
            return 'Helvetica'
        
    def generate_label_pdf(self, data, output_path: str):
        """
        生成标签PDF
        
        Args:
            data: 数据字典，包含客户编号、主题等信息
            output_path: 输出文件路径
        """
        try:
            c = canvas.Canvas(output_path, pagesize=self.page_size)
            width, height = self.page_size
            
            # 设置标题
            c.setFont(self.chinese_font, 24)
            title = "多级标签PDF"
            title_width = c.stringWidth(title, self.chinese_font, 24)
            c.drawString((width - title_width) / 2, height - 60, title)
            
            # 设置内容区域
            y_position = height - 120
            
            # 绘制数据信息
            c.setFont(self.chinese_font, 14)
            
            # 客户编号
            if 'customer_code' in data:
                c.drawString(self.margin, y_position, f"客户编号: {data['customer_code']}")
                y_position -= 30
            
            # 主题
            if 'subject' in data:
                c.drawString(self.margin, y_position, f"主题: {data['subject']}")
                y_position -= 30
            
            # 其他字段
            for key, value in data.items():
                if key not in ['customer_code', 'subject']:
                    c.drawString(self.margin, y_position, f"{key}: {value}")
                    y_position -= 25
                    if y_position < 50:  # 如果接近页面底部，创建新页面
                        c.showPage()
                        y_position = height - 60
            
            c.save()
            return True
        except Exception as e:
            raise Exception(f"生成PDF失败: {str(e)}")
    
    def batch_generate(self, data_list, output_dir: str):
        """
        批量生成PDF
        
        Args:
            data_list: 数据列表
            output_dir: 输出目录
        """
        try:
            output_path = Path(output_dir)
            output_path.mkdir(parents=True, exist_ok=True)
            
            generated_files = []
            
            for i, data in enumerate(data_list):
                filename = f"label_{i+1}.pdf"
                if 'customer_code' in data:
                    filename = f"label_{data['customer_code']}.pdf"
                
                file_path = output_path / filename
                self.generate_label_pdf(data, str(file_path))
                generated_files.append(str(file_path))
            
            return generated_files
        except Exception as e:
            raise Exception(f"批量生成PDF失败: {str(e)}")
    
    def generate_multi_label_pdf(self, data_list, output_path: str):
        """
        生成包含多个标签的单个PDF文件
        
        Args:
            data_list: 数据列表
            output_path: 输出文件路径
        """
        try:
            c = canvas.Canvas(output_path, pagesize=self.page_size)
            width, height = self.page_size
            
            # 设置标题
            c.setFont(self.chinese_font, 24)
            title = "多级标签PDF批量生成"
            title_width = c.stringWidth(title, self.chinese_font, 24)
            c.drawString((width - title_width) / 2, height - 60, title)
            
            y_position = height - 120
            
            for i, data in enumerate(data_list):
                # 检查是否需要新页面
                if y_position < 100:
                    c.showPage()
                    y_position = height - 60
                
                # 绘制分隔线
                if i > 0:
                    c.setStrokeColor(black)
                    c.line(self.margin, y_position + 10, width - self.margin, y_position + 10)
                    y_position -= 20
                
                # 绘制标签编号
                c.setFont(self.chinese_font, 16)
                c.drawString(self.margin, y_position, f"标签 #{i+1}")
                y_position -= 25
                
                # 绘制数据
                c.setFont(self.chinese_font, 12)
                
                for key, value in data.items():
                    if y_position < 50:
                        c.showPage()
                        y_position = height - 60
                    
                    c.drawString(self.margin + 20, y_position, f"{key}: {value}")
                    y_position -= 20
                
                y_position -= 10  # 标签间距
            
            c.save()
            return True
        except Exception as e:
            raise Exception(f"生成多标签PDF失败: {str(e)}")