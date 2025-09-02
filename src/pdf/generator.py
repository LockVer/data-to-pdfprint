"""
PDF生成器

使用ReportLab生成PDF文档
"""

from reportlab.lib.pagesizes import A4, LETTER, A5
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.colors import CMYKColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pathlib import Path
from typing import List, Dict, Any
import math
import os
import platform


class PDFGenerator:
    """
    PDF生成器类
    """

    def __init__(self, max_pages_per_file: int = 100):
        """
        初始化PDF生成器
        
        Args:
            max_pages_per_file: 每个PDF文件的最大页数限制
        """
        self.page_size = (90 * mm, 50 * mm)  # 90mm x 50mm标签尺寸
        self.margin = 2 * mm
        self.font_name = "Helvetica"
        self.font_size = 8
        self.chinese_font_name = "SimHei"
        self.max_pages_per_file = max_pages_per_file
        self._register_chinese_font()

    def set_page_size(self, size: str):
        """
        设置页面大小

        Args:
            size: 页面大小 ('A4', 'LETTER', 'A5')
        """
        size_map = {"A4": A4, "LETTER": LETTER, "A5": A5}

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
            if template and hasattr(template, "render"):
                template.render(c, data, width, height)
            else:
                # 使用默认布局
                self._render_default_layout(c, data, width, height)

            # 保存PDF
            c.save()

        except Exception as e:
            raise Exception(f"生成PDF失败: {e}")

    def batch_generate(
        self, template, data_list: List[Dict[str, Any]], output_dir: str
    ) -> List[str]:
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
            if "name" in data:
                filename = f"label_{data['name']}_{i+1:03d}.pdf"

            file_path = output_path / filename

            try:
                self.generate_from_template(template, data, str(file_path))
                output_files.append(str(file_path))
            except Exception as e:
                print(f"生成第{i+1}个PDF失败: {e}")
                continue

        return output_files

    def create_multi_level_pdfs(
        self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str
    ) -> Dict[str, str]:
        """
        创建多级标签PDF

        Args:
            data: Excel数据
            params: 用户参数 (张/盒, 盒/小箱, 小箱/大箱, 选择外观)
            output_dir: 输出目录

        Returns:
            生成的文件路径字典
        """
        # 计算数量 - 三级结构：张→盒→小箱→大箱
        total_pieces = int(data["总张数"])
        pieces_per_box = int(params["张/盒"])  # 用户输入
        boxes_per_small_box = int(params["盒/小箱"])  # 用户输入
        small_boxes_per_large_box = int(params["小箱/大箱"])  # 用户输入

        # 计算各级数量 (使用向上取整处理余数)
        total_boxes = math.ceil(total_pieces / pieces_per_box)  # 总盒数
        total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)  # 总小箱数
        total_large_boxes = math.ceil(
            total_small_boxes / small_boxes_per_large_box
        )  # 总大箱数

        # 计算余数信息 - 基于新的三级结构
        remaining_pieces_in_last_box = total_pieces % pieces_per_box
        remaining_boxes_in_last_small_box = total_boxes % boxes_per_small_box
        remaining_small_boxes_in_last_large_box = (
            total_small_boxes % small_boxes_per_large_box
        )

        # 记录余数信息
        remainder_info = {
            "last_box_pieces": (
                remaining_pieces_in_last_box
                if remaining_pieces_in_last_box > 0
                else pieces_per_box
            ),
            "last_small_box_boxes": (
                remaining_boxes_in_last_small_box
                if remaining_boxes_in_last_small_box > 0
                else boxes_per_small_box
            ),
            "last_large_box_small_boxes": (
                remaining_small_boxes_in_last_large_box
                if remaining_small_boxes_in_last_large_box > 0
                else small_boxes_per_large_box
            ),
            "total_boxes": total_boxes,
        }

        # 创建输出目录 - 清理文件名中的特殊字符
        clean_theme = data['主题'].replace('\n', ' ').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('*', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
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

        self._create_box_label(data, params, str(box_label_path), selected_appearance)
        generated_files["盒标"] = str(box_label_path)

        # 生成小箱标
        small_box_path = (
            full_output_dir / f"{data['客户编码']}+{clean_theme}+小箱标.pdf"
        )
        self._create_small_box_label(
            data, params, str(small_box_path), total_small_boxes, remainder_info
        )
        generated_files["小箱标"] = str(small_box_path)

        # 生成大箱标
        large_box_path = (
            full_output_dir / f"{data['客户编码']}+{clean_theme}+大箱标.pdf"
        )
        self._create_large_box_label(
            data,
            params,
            str(large_box_path),
            total_large_boxes,
            total_small_boxes,
            remainder_info,
        )
        generated_files["大箱标"] = str(large_box_path)

        return generated_files

    def _create_box_label(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, style: str
    ):
        """创建盒标 - 支持分页限制的多页PDF"""
        # 计算总盒数
        total_pieces = int(data["总张数"])
        pieces_per_box = int(params["张/盒"])
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 获取盒标内容 - 从Excel的H11和B11单元格
        import pandas as pd
        df = pd.read_excel('/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx', header=None)
        top_text = str(df.iloc[10, 7])  # H11单元格
        base_number = str(df.iloc[10, 1])  # B11单元格 (TAG01001)
        
        # 计算需要的文件数量
        total_files = math.ceil(total_boxes / self.max_pages_per_file)
        
        # 分页生成多个PDF文件
        for file_num in range(total_files):
            # 计算当前文件的页数范围
            start_box = file_num * self.max_pages_per_file + 1
            end_box = min((file_num + 1) * self.max_pages_per_file, total_boxes)
            
            # 生成文件名
            if total_files == 1:
                current_output_path = output_path
            else:
                path_obj = Path(output_path)
                current_output_path = str(path_obj.parent / f"{path_obj.stem}_part{file_num + 1:02d}{path_obj.suffix}")
            
            # 创建当前文件的PDF
            self._create_single_box_label_file(
                data, params, current_output_path, style, 
                start_box, end_box, top_text, base_number
            )

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
        c.setSubject("Product Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)
        
        # 真正的三等分留白布局：每个留白区域高度相等
        blank_height = height / 5  # 每个留白区域高度：10mm
        
        # 布局位置计算（确保三个留白区域等高）
        # 留白1: height到(height - blank_height) = 40mm到50mm
        # 产品名称区域: (height - blank_height)到(height - 2*blank_height) = 30mm到40mm  
        # 留白2: (height - 2*blank_height)到(height - 3*blank_height) = 20mm到30mm
        # 序列号区域: (height - 3*blank_height)到(height - 4*blank_height) = 10mm到20mm
        # 留白3: (height - 4*blank_height)到0 = 0mm到10mm
        
        top_text_y = height - 1.5 * blank_height      # 产品名称居中在区域2
        serial_number_y = height - 3.5 * blank_height # 序列号居中在区域4

        # 生成指定范围的盒标
        for box_num in range(start_box, end_box + 1):
            if box_num > start_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 生成当前盒标的编号 - 基于B11单元格的格式
            # B11是LAN01001，提取前缀和数字部分
            import re
            match = re.match(r'([A-Z]+)(\d+)', base_number)
            if match:
                prefix = match.group(1)  # LAN
                base_num = int(match.group(2))  # 01001
                current_number = f"{prefix}{base_num + box_num - 1:05d}"
            else:
                # 备用方案
                current_number = f"DSK{box_num:05d}"
            
            # 根据外观选择不同的样式
            if style == "外观一":
                self._render_appearance_one(c, width, top_text, current_number, top_text_y, serial_number_y)
            else:
                # 外观二需要额外的票数信息
                ticket_count = params["张/盒"]
                self._render_appearance_two(c, width, top_text, ticket_count, current_number, top_text_y, serial_number_y)

        c.save()

    def _render_appearance_one(self, c, width, top_text, serial_number, top_text_y, serial_number_y):
        """渲染外观一：简洁标准样式"""
        # 使用多次绘制实现加粗效果
        c.setFillColor(CMYKColor(0, 0, 0, 1))
        c.setFont("Helvetica-Bold", 22)
        
        # 上部文本 - 多次绘制增加粗细
        for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
            c.drawCentredString(width / 2 + offset[0], top_text_y + offset[1], top_text)

        # 下部序列号 - 多次绘制增加粗细  
        for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
            c.drawCentredString(width / 2 + offset[0], serial_number_y + offset[1], serial_number)

    def _render_appearance_two(self, c, width, game_title, ticket_count, serial_number, top_y, bottom_y):
        """渲染外观二：三行信息格式"""
        c.setFillColor(CMYKColor(0, 0, 0, 1))
        c.setFont("Helvetica-Bold", 12)
        
        # 获取页面高度
        page_height = self.page_size[1]
        
        # 左边距 - 统一的左边距
        left_margin = 4 * mm
        
        # Game title: 距离顶部更大的距离
        game_title_y = page_height - 12 * mm  # 距离顶部12mm
        c.drawString(left_margin, game_title_y, f"Game title: {game_title}")
        
        # Ticket count: 左下区域，增加与Serial的间距
        ticket_count_y = 15 * mm  # 距离底部15mm
        c.drawString(left_margin, ticket_count_y, f"Ticket count: {ticket_count}")
        
        # Serial: 距离底部
        serial_y = 6 * mm  # 距离底部6mm
        c.drawString(left_margin, serial_y, f"Serial: {serial_number}")

    def _create_small_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        total_small_boxes: int,
        remainder_info: Dict[str, Any],
    ):
        """创建小箱标 - 支持分页限制的多页PDF"""
        # 获取Excel数据
        import pandas as pd
        df = pd.read_excel('/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx', header=None)
        theme_text = str(df.iloc[10, 7])  # H11单元格
        base_number = str(df.iloc[10, 1])  # B11单元格
        remark_text = str(df.iloc[3, 0])   # A4单元格
        
        # 计算参数
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        pieces_per_small_box = pieces_per_box * boxes_per_small_box
        
        # 计算需要的文件数量
        total_files = math.ceil(total_small_boxes / self.max_pages_per_file)
        
        # 分页生成多个PDF文件
        for file_num in range(total_files):
            # 计算当前文件的页数范围
            start_small_box = file_num * self.max_pages_per_file + 1
            end_small_box = min((file_num + 1) * self.max_pages_per_file, total_small_boxes)
            
            # 生成文件名
            if total_files == 1:
                current_output_path = output_path
            else:
                path_obj = Path(output_path)
                current_output_path = str(path_obj.parent / f"{path_obj.stem}_part{file_num + 1:02d}{path_obj.suffix}")
            
            # 创建当前文件的PDF
            self._create_single_small_box_label_file(
                data, params, current_output_path, start_small_box, end_small_box,
                theme_text, base_number, remark_text, pieces_per_small_box, 
                boxes_per_small_box, total_small_boxes
            )

    def _create_single_small_box_label_file(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
        start_small_box: int, end_small_box: int, theme_text: str, base_number: str,
        remark_text: str, pieces_per_small_box: int, boxes_per_small_box: int, total_small_boxes: int
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

            # 计算当前小箱的序列号范围
            import re
            match = re.match(r'([A-Z]+)(\d+)', base_number)
            if match:
                prefix = match.group(1)  # LAN
                base_num = int(match.group(2))  # 01001
                
                # 计算序列号范围
                start_serial = base_num + (small_box_num - 1) * boxes_per_small_box
                end_serial = start_serial + boxes_per_small_box - 1
                
                if boxes_per_small_box == 1:
                    # 一盒一小箱：序列号相同
                    serial_range = f"{prefix}{start_serial:05d}-{prefix}{start_serial:05d}"
                else:
                    # 多盒一小箱：序列号范围
                    serial_range = f"{prefix}{start_serial:05d}-{prefix}{end_serial:05d}"
            else:
                serial_range = f"DSK{small_box_num:05d}-DSK{small_box_num:05d}"

            # 绘制表格
            self._draw_small_box_table(c, width, height, theme_text, pieces_per_small_box, 
                                     serial_range, small_box_num, total_small_boxes, remark_text)

        c.save()

    def _draw_small_box_table(self, c, width, height, theme_text, pieces_per_small_box, 
                            serial_range, small_box_num, total_small_boxes, remark_text):
        """绘制小箱标表格"""
        # 表格尺寸和位置
        table_width = width - 4 * mm
        table_height = height - 4 * mm
        table_x = 2 * mm
        table_y = 2 * mm
        
        # 高度分配：Quantity行占2/6，其他4行各占1/6
        base_row_height = table_height / 6
        quantity_row_height = base_row_height * 2  # Quantity行双倍高度
        
        # 列宽 (标签列:数据列 = 1:2)
        label_col_width = table_width / 3
        data_col_width = table_width * 2 / 3
        
        # 绘制表格边框
        c.setStrokeColor(CMYKColor(0, 0, 0, 1))
        c.setLineWidth(1)
        c.rect(table_x, table_y, table_width, table_height)
        
        # 计算各行的Y坐标
        row_positions = []
        current_y = table_y
        # 从底部开始：Remark, Carton No, Quantity(双倍), Theme, Item
        for height in [base_row_height, base_row_height, quantity_row_height, base_row_height, base_row_height]:
            row_positions.append(current_y)
            current_y += height
        
        # 绘制行线
        for i in range(1, 5):
            y = row_positions[i]
            c.line(table_x, y, table_x + table_width, y)
        
        # 绘制列线
        col_x = table_x + label_col_width
        c.line(col_x, table_y, col_x, table_y + table_height)
        
        # 绘制Quantity行的分隔线（上层和下层之间）
        quantity_split_y = row_positions[2] + quantity_row_height / 2
        c.line(col_x, quantity_split_y, table_x + table_width, quantity_split_y)
        
        # 表格内容
        c.setFont("Helvetica-Bold", 10)
        
        # 计算居中位置
        label_center_x = table_x + label_col_width / 2  # 标签列居中
        data_center_x = col_x + data_col_width / 2      # 数据列居中
        
        # 行1: Item (第5行，从上往下)
        item_y = row_positions[4] + base_row_height/2
        c.drawCentredString(label_center_x, item_y, "Item:")
        c.drawCentredString(data_center_x, item_y, "Paper Cards")
        
        # 行2: Theme (第4行)
        theme_y = row_positions[3] + base_row_height/2
        c.drawCentredString(label_center_x, theme_y, "Theme:")
        c.drawCentredString(data_center_x, theme_y, theme_text)
        
        # 行3: Quantity (第3行，双倍高度)
        quantity_label_y = row_positions[2] + quantity_row_height/2
        c.drawCentredString(label_center_x, quantity_label_y, "Quantity:")
        # 上层：票数（在分隔线上方居中）
        upper_y = row_positions[2] + quantity_row_height * 3/4
        c.drawCentredString(data_center_x, upper_y, f"{pieces_per_small_box}PCS")
        # 下层：序列号（在分隔线下方居中）
        lower_y = row_positions[2] + quantity_row_height/4
        c.drawCentredString(data_center_x, lower_y, serial_range)
        
        # 行4: Carton No (第2行)
        carton_y = row_positions[1] + base_row_height/2
        c.drawCentredString(label_center_x, carton_y, "Carton No:")
        c.drawCentredString(data_center_x, carton_y, f"{small_box_num}/{total_small_boxes}")
        
        # 行5: Remark (第1行)
        remark_y = row_positions[0] + base_row_height/2
        c.drawCentredString(label_center_x, remark_y, "Remark:")
        c.drawCentredString(data_center_x, remark_y, remark_text)

    def _create_large_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        total_large_boxes: int,
        total_small_boxes: int,
        remainder_info: Dict[str, Any],
    ):
        """创建大箱标 - 支持分页限制的多页PDF"""
        # 获取Excel数据
        import pandas as pd
        df = pd.read_excel('/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx', header=None)
        theme_text = str(df.iloc[10, 7])  # H11单元格
        base_number = str(df.iloc[10, 1])  # B11单元格
        remark_text = str(df.iloc[3, 0])   # A4单元格
        
        # 计算参数
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        small_boxes_per_large_box = int(params["小箱/大箱"])
        pieces_per_large_box = pieces_per_box * boxes_per_small_box * small_boxes_per_large_box
        
        # 计算需要的文件数量
        total_files = math.ceil(total_large_boxes / self.max_pages_per_file)
        
        # 分页生成多个PDF文件
        for file_num in range(total_files):
            # 计算当前文件的页数范围
            start_large_box = file_num * self.max_pages_per_file + 1
            end_large_box = min((file_num + 1) * self.max_pages_per_file, total_large_boxes)
            
            # 生成文件名
            if total_files == 1:
                current_output_path = output_path
            else:
                path_obj = Path(output_path)
                current_output_path = str(path_obj.parent / f"{path_obj.stem}_part{file_num + 1:02d}{path_obj.suffix}")
            
            # 创建当前文件的PDF
            self._create_single_large_box_label_file(
                data, params, current_output_path, start_large_box, end_large_box,
                theme_text, base_number, remark_text, pieces_per_large_box,
                boxes_per_small_box, small_boxes_per_large_box, total_large_boxes
            )

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

            # 计算当前大箱包含的盒序列号范围
            import re
            match = re.match(r'([A-Z]+)(\d+)', base_number)
            if match:
                prefix = match.group(1)  # LAN
                base_num = int(match.group(2))  # 01001
                
                # 计算大箱内盒序列号范围
                # 每大箱包含的盒数 = 盒/小箱 × 小箱/大箱
                boxes_per_large_box = boxes_per_small_box * small_boxes_per_large_box
                start_box_serial = base_num + (large_box_num - 1) * boxes_per_large_box
                end_box_serial = start_box_serial + boxes_per_large_box - 1
                
                box_serial_range = f"{prefix}{start_box_serial:05d}-{prefix}{end_box_serial:05d}"
            else:
                box_serial_range = f"DSK{large_box_num:05d}-DSK{large_box_num:05d}"

            # 绘制表格
            self._draw_large_box_table(c, width, height, theme_text, pieces_per_large_box,
                                     box_serial_range, large_box_num, total_large_boxes, remark_text)

        c.save()

    def _draw_large_box_table(self, c, width, height, theme_text, pieces_per_large_box,
                            box_serial_range, large_box_num, total_large_boxes, remark_text):
        """绘制大箱标表格"""
        # 复用小箱标的表格绘制逻辑，只是数据不同
        # 表格尺寸和位置
        table_width = width - 4 * mm
        table_height = height - 4 * mm
        table_x = 2 * mm
        table_y = 2 * mm
        
        # 高度分配：Quantity行占2/6，其他4行各占1/6
        base_row_height = table_height / 6
        quantity_row_height = base_row_height * 2
        
        # 列宽 (标签列:数据列 = 1:2)
        label_col_width = table_width / 3
        
        # 绘制表格边框
        c.setStrokeColor(CMYKColor(0, 0, 0, 1))
        c.setLineWidth(1)
        c.rect(table_x, table_y, table_width, table_height)
        
        # 计算各行的Y坐标
        row_positions = []
        current_y = table_y
        # 从底部开始：Remark, Carton No, Quantity(双倍), Theme, Item
        for height_val in [base_row_height, base_row_height, quantity_row_height, base_row_height, base_row_height]:
            row_positions.append(current_y)
            current_y += height_val
        
        # 绘制行线
        for i in range(1, 5):
            y = row_positions[i]
            c.line(table_x, y, table_x + table_width, y)
        
        # 绘制列线
        col_x = table_x + label_col_width
        c.line(col_x, table_y, col_x, table_y + table_height)
        
        # 绘制Quantity行的分隔线
        quantity_split_y = row_positions[2] + quantity_row_height / 2
        c.line(col_x, quantity_split_y, table_x + table_width, quantity_split_y)
        
        # 表格内容
        c.setFont("Helvetica-Bold", 10)
        
        # 计算居中位置
        label_center_x = table_x + label_col_width / 2
        data_center_x = col_x + (table_width - label_col_width) / 2
        
        # 行1: Item
        item_y = row_positions[4] + base_row_height/2
        c.drawCentredString(label_center_x, item_y, "Item:")
        c.drawCentredString(data_center_x, item_y, "Paper Cards")
        
        # 行2: Theme
        theme_y = row_positions[3] + base_row_height/2
        c.drawCentredString(label_center_x, theme_y, "Theme:")
        c.drawCentredString(data_center_x, theme_y, theme_text)
        
        # 行3: Quantity (双倍高度)
        quantity_label_y = row_positions[2] + quantity_row_height/2
        c.drawCentredString(label_center_x, quantity_label_y, "Quantity:")
        # 上层：每大箱票数
        upper_y = row_positions[2] + quantity_row_height * 3/4
        c.drawCentredString(data_center_x, upper_y, f"{pieces_per_large_box}PCS")
        # 下层：大箱内盒序列号范围
        lower_y = row_positions[2] + quantity_row_height/4
        c.drawCentredString(data_center_x, lower_y, box_serial_range)
        
        # 行4: Carton No
        carton_y = row_positions[1] + base_row_height/2
        c.drawCentredString(label_center_x, carton_y, "Carton No:")
        c.drawCentredString(data_center_x, carton_y, f"{large_box_num}/{total_large_boxes}")
        
        # 行5: Remark
        remark_y = row_positions[0] + base_row_height/2
        c.drawCentredString(label_center_x, remark_y, "Remark:")
        c.drawCentredString(data_center_x, remark_y, remark_text)

    def create_label_pdf(self, data: Dict[str, Any], output_path: str):
        """
        创建标签PDF (保持兼容性)

        Args:
            data: 标签数据
            output_path: 输出路径
        """
        self.generate_from_template(None, data, output_path)

    def _render_default_layout(
        self, canvas_obj, data: Dict[str, Any], width: float, height: float
    ):
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
                    "/Library/Fonts/SimHei.ttf",
                ]
            elif system == "Windows":
                font_paths = [
                    "C:/Windows/Fonts/simhei.ttf",
                    "C:/Windows/Fonts/simsun.ttc",
                    "C:/Windows/Fonts/msyh.ttc",
                ]
            elif system == "Linux":
                font_paths = [
                    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
                    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
                ]

            # 尝试注册第一个可用的字体
            for font_path in font_paths:
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(
                            TTFont(self.chinese_font_name, font_path)
                        )
                        self.font_name = self.chinese_font_name
                        return
                    except Exception:
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
            if "\u4e00" <= char <= "\u9fff":
                return True
        return False
