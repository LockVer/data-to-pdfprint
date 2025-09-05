"""
常规模板 - 标准的多级标签PDF生成
"""
import math
from pathlib import Path
from typing import Dict, Any
from reportlab.pdfgen import canvas
from reportlab.lib.colors import CMYKColor
from reportlab.lib.units import mm

# 导入基础工具类
from src.utils.pdf_base import PDFBaseUtils
from src.utils.font_manager import font_manager
from src.utils.text_processor import text_processor
from src.utils.excel_data_extractor import ExcelDataExtractor

# 导入常规模板专属数据处理器和渲染器
from src.pdf.regular.data_processor import regular_data_processor
from src.pdf.regular.renderer import regular_renderer


class RegularTemplate(PDFBaseUtils):
    """常规模板处理类"""
    
    def __init__(self, max_pages_per_file: int = 100):
        """初始化常规模板"""
        super().__init__(max_pages_per_file)
    
    def create_multi_level_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """
        创建常规模板的多级标签PDF

        Args:
            data: Excel数据
            params: 用户参数 (张/盒, 盒/小箱, 小箱/大箱, 选择外观)
            output_dir: 输出目录

        Returns:
            生成的文件路径字典
        """
        # 计算数量 - 三级结构：张→盒→小箱→大箱
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        small_boxes_per_large_box = int(params["小箱/大箱"])

        # 计算各级数量
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)
        total_large_boxes = math.ceil(total_small_boxes / small_boxes_per_large_box)

        # 计算余数信息
        remaining_pieces_in_last_box = total_pieces % pieces_per_box
        remaining_boxes_in_last_small_box = total_boxes % boxes_per_small_box
        remaining_small_boxes_in_last_large_box = total_small_boxes % small_boxes_per_large_box

        remainder_info = {
            "total_boxes": total_boxes,
            "remaining_pieces_in_last_box": (
                pieces_per_box if remaining_pieces_in_last_box == 0 else remaining_pieces_in_last_box
            ),
            "remaining_boxes_in_last_small_box": (
                boxes_per_small_box if remaining_boxes_in_last_small_box == 0 else remaining_boxes_in_last_small_box
            ),
            "remaining_small_boxes_in_last_large_box": (
                small_boxes_per_large_box if remaining_small_boxes_in_last_large_box == 0 else remaining_small_boxes_in_last_large_box
            ),
        }

        # 创建输出目录
        clean_theme = data['主题'].replace('\n', ' ').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('*', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').replace('!', '_')
        folder_name = f"{data['客户编码']}+{clean_theme}+标签"
        full_output_dir = Path(output_dir) / folder_name
        full_output_dir.mkdir(parents=True, exist_ok=True)

        generated_files = {}

        # 生成盒标 (只生成用户选择的外观)
        selected_appearance = params["选择外观"]
        box_label_path = (
            full_output_dir
            / f"{data['客户编码']}+{clean_theme}+盒标+{selected_appearance}.pdf"
        )

        self._create_box_label(data, params, str(box_label_path), selected_appearance, excel_file_path)
        generated_files["盒标"] = str(box_label_path)

        # 生成小箱标
        small_box_path = (
            full_output_dir / f"{data['客户编码']}+{clean_theme}+小箱标.pdf"
        )
        self._create_small_box_label(
            data, params, str(small_box_path), total_small_boxes, remainder_info, excel_file_path
        )
        generated_files["小箱标"] = str(small_box_path)

        # 生成大箱标
        large_box_path = (
            full_output_dir / f"{data['客户编码']}+{clean_theme}+大箱标.pdf"
        )
        self._create_large_box_label(
            data, params, str(large_box_path), total_large_boxes, total_small_boxes, remainder_info, excel_file_path
        )
        generated_files["大箱标"] = str(large_box_path)

        return generated_files

    def _create_box_label(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, style: str, excel_file_path: str = None):
        """创建盒标 - 支持分页限制的多页PDF"""
        # 计算总盒数
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 获取盒标内容 - 优先使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        # 使用常规模板专属数据处理器
        excel_data = regular_data_processor.extract_box_label_data(excel_path)
        top_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DSK00001'
        
        # 分页生成PDF
        current_page = 1
        boxes_processed = 0
        
        while boxes_processed < total_boxes:
            remaining_boxes = total_boxes - boxes_processed
            boxes_in_current_file = min(self.max_pages_per_file, remaining_boxes)
            
            # 构建文件名
            if total_boxes <= self.max_pages_per_file:
                current_output_path = output_path
            else:
                base_path = Path(output_path)
                current_output_path = (
                    base_path.parent / f"{base_path.stem}_第{current_page}部分{base_path.suffix}"
                )
            
            # 创建当前文件
            self._create_single_box_label_file(
                data, params, str(current_output_path), style,
                boxes_processed + 1, boxes_processed + boxes_in_current_file,
                top_text, base_number
            )
            
            boxes_processed += boxes_in_current_file
            current_page += 1

    def _create_single_box_label_file(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, 
        style: str, start_box: int, end_box: int, top_text: str, base_number: str
    ):
        """创建单个盒标PDF文件"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"盒标-{style}-{start_box}到{end_box}")
        c.setSubject("Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)
        
        # 真正的三等分留白布局：每个留白区域高度相等
        blank_height = height / 5  # 每个留白区域高度：10mm
        
        # 布局位置计算（确保三个留白区域等高）
        top_text_y = height - 1.5 * blank_height      # 产品名称居中在区域2
        serial_number_y = height - 3.5 * blank_height # 序列号居中在区域4

        # 生成指定范围的盒标
        for box_num in range(start_box, end_box + 1):
            if box_num > start_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 解析基础序列号格式
            import re
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取数字前的前缀和数字
                digit_start = match.start()
                prefix = base_number[:digit_start]
                base_num = int(match.group(1))
                
                # 计算当前序列号
                current_num = base_num + (box_num - 1)
                current_number = f"{prefix}{current_num:05d}"
            else:
                # 如果无法解析，使用简单递增
                current_number = f"BOX{box_num:05d}"

            # 根据选择的外观渲染
            if style == "外观一":
                regular_renderer.render_appearance_one(c, width, top_text, current_number, top_text_y, serial_number_y)
            else:
                # 获取票数信息用于外观二
                total_pieces = int(float(data["总张数"]))
                pieces_per_box = int(params["张/盒"])
                regular_renderer.render_appearance_two(c, width, self.page_size, top_text, pieces_per_box, current_number, top_text_y, serial_number_y)

        c.save()



    def _create_small_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        total_small_boxes: int,
        remainder_info: Dict[str, Any],
        excel_file_path: str = None,
    ):
        """创建小箱标"""
        # 获取Excel数据 - 使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用常规模板专属数据处理器
        excel_data = regular_data_processor.extract_small_box_label_data(excel_path)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        remark_text = excel_data.get('客户编码') or 'Unknown Client'
        
        # 计算参数
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        pieces_per_small_box = pieces_per_box * boxes_per_small_box
        
        # 分页生成PDF
        current_page = 1
        small_boxes_processed = 0
        
        while small_boxes_processed < total_small_boxes:
            remaining_small_boxes = total_small_boxes - small_boxes_processed
            small_boxes_in_current_file = min(self.max_pages_per_file, remaining_small_boxes)
            
            # 构建文件名
            if total_small_boxes <= self.max_pages_per_file:
                current_output_path = output_path
            else:
                base_path = Path(output_path)
                current_output_path = (
                    base_path.parent / f"{base_path.stem}_第{current_page}部分{base_path.suffix}"
                )
            
            # 创建当前文件
            self._create_single_small_box_label_file(
                data, params, str(current_output_path),
                small_boxes_processed + 1, small_boxes_processed + small_boxes_in_current_file,
                theme_text, base_number, remark_text, pieces_per_small_box, 
                boxes_per_small_box, total_small_boxes
            )
            
            small_boxes_processed += small_boxes_in_current_file
            current_page += 1

    def _create_single_small_box_label_file(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
        start_small_box: int, end_small_box: int, theme_text: str, base_number: str,
        remark_text: str, pieces_per_small_box: int, boxes_per_small_box: int, 
        total_small_boxes: int
    ):
        """创建单个小箱标PDF文件"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"小箱标-{start_small_box}到{end_small_box}")
        c.setSubject("Small Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成指定范围的小箱标
        for small_box_num in range(start_small_box, end_small_box + 1):
            if small_box_num > start_small_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 计算序列号范围
            import re
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取数字前的前缀和基础数字
                digit_start = match.start()
                prefix = base_number[:digit_start]
                base_num = int(match.group(1))
                
                # 计算当前小箱的序列号范围
                start_box_num = (small_box_num - 1) * boxes_per_small_box + 1
                end_box_num = small_box_num * boxes_per_small_box
                
                start_serial_num = base_num + (start_box_num - 1)
                end_serial_num = base_num + (end_box_num - 1)
                
                start_serial = f"{prefix}{start_serial_num:05d}"
                end_serial = f"{prefix}{end_serial_num:05d}"
                serial_range = f"{start_serial}-{end_serial}"
            else:
                serial_range = f"BOX{small_box_num:05d}"

            # 绘制小箱标表格
            regular_renderer.draw_small_box_table(c, width, height, theme_text, pieces_per_small_box, 
                                                 serial_range, str(small_box_num), remark_text)

        c.save()


    def _create_large_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        total_large_boxes: int,
        total_small_boxes: int,
        remainder_info: Dict[str, Any],
        excel_file_path: str = None,
    ):
        """创建大箱标"""
        # 获取Excel数据 - 使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用常规模板专属数据处理器
        excel_data = regular_data_processor.extract_large_box_label_data(excel_path)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        remark_text = excel_data.get('客户编码') or 'Unknown Client'
        
        # 计算参数 - 大箱标专用
        pieces_per_box = int(params["张/盒"])  
        boxes_per_small_box = int(params["盒/小箱"]) 
        small_boxes_per_large_box = int(params["小箱/大箱"])  
        
        pieces_per_large_box = pieces_per_box * boxes_per_small_box * small_boxes_per_large_box
        
        # 分页生成PDF
        current_page = 1
        large_boxes_processed = 0
        
        while large_boxes_processed < total_large_boxes:
            remaining_large_boxes = total_large_boxes - large_boxes_processed
            large_boxes_in_current_file = min(self.max_pages_per_file, remaining_large_boxes)
            
            # 构建文件名
            if total_large_boxes <= self.max_pages_per_file:
                current_output_path = output_path
            else:
                base_path = Path(output_path)
                current_output_path = (
                    base_path.parent / f"{base_path.stem}_第{current_page}部分{base_path.suffix}"
                )
            
            # 创建当前文件
            self._create_single_large_box_label_file(
                data, params, str(current_output_path),
                large_boxes_processed + 1, large_boxes_processed + large_boxes_in_current_file,
                theme_text, base_number, remark_text, pieces_per_large_box, 
                boxes_per_small_box, small_boxes_per_large_box, total_large_boxes
            )
            
            large_boxes_processed += large_boxes_in_current_file
            current_page += 1

    def _create_single_large_box_label_file(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
        start_large_box: int, end_large_box: int, theme_text: str, base_number: str,
        remark_text: str, pieces_per_large_box: int, boxes_per_small_box: int, 
        small_boxes_per_large_box: int, total_large_boxes: int
    ):
        """创建单个大箱标PDF文件"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"大箱标-{start_large_box}到{end_large_box}")
        c.setSubject("Large Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成指定范围的大箱标
        for large_box_num in range(start_large_box, end_large_box + 1):
            if large_box_num > start_large_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 计算序列号范围
            import re
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取数字前的前缀和基础数字
                digit_start = match.start()
                prefix = base_number[:digit_start]
                base_num = int(match.group(1))
                
                # 计算当前大箱包含的盒标范围
                boxes_per_large_box = boxes_per_small_box * small_boxes_per_large_box
                start_box_num = (large_box_num - 1) * boxes_per_large_box + 1
                end_box_num = large_box_num * boxes_per_large_box
                
                start_serial_num = base_num + (start_box_num - 1)
                end_serial_num = base_num + (end_box_num - 1)
                
                start_serial = f"{prefix}{start_serial_num:05d}"
                end_serial = f"{prefix}{end_serial_num:05d}"
                serial_range = f"{start_serial}-{end_serial}"
            else:
                serial_range = f"LARGE{large_box_num:05d}"

            # 绘制大箱标表格
            regular_renderer.draw_large_box_table(c, width, height, theme_text, pieces_per_large_box,
                                                 serial_range, str(large_box_num), remark_text)

        c.save()

