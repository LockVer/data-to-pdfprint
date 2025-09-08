"""
套盒模板渲染器
负责套盒模板的所有PDF绘制和渲染逻辑
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import CMYKColor
from reportlab.lib.units import mm

# 导入工具类
from src.utils.font_manager import font_manager
from src.utils.text_processor import text_processor


class NestedBoxRenderer:
    """套盒模板渲染器 - 专门负责套盒模板的PDF绘制逻辑"""
    
    def __init__(self):
        """初始化套盒渲染器"""
        self.renderer_type = "nested_box"
    
    def render_nested_appearance_one(self, c, width, top_text, current_number, top_text_y, serial_number_y):
        """套盒模板盒标外观一渲染"""
        clean_top_text = text_processor.clean_text_for_font(top_text)
        font_manager.set_best_font(c, 14, bold=True)
        
        # 绘制Game title和序列号 - 加粗效果
        for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
            c.drawCentredString(width / 2 + offset[0], top_text_y + offset[1], clean_top_text)
            c.drawCentredString(width / 2 + offset[0], serial_number_y + offset[1], current_number)

    def render_nested_appearance_two(self, c, width, top_text, current_number, top_text_y, serial_number_y):
        """套盒模板盒标外观二渲染"""
        clean_top_text = text_processor.clean_text_for_font(top_text)
        font_manager.set_best_font(c, 14, bold=True)
        
        # 外观二：Game title左对齐，但溢出文本居中
        max_width = width * 0.8
        title_lines = text_processor.wrap_text_to_fit(c, clean_top_text, max_width, font_manager.get_chinese_font_name(), 14)
        
        if len(title_lines) > 1:
            # 首行左对齐，其他行居中
            for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
                c.drawString(width * 0.1 + offset[0], top_text_y + 15 + offset[1], title_lines[0])
            for i, line in enumerate(title_lines[1:], 1):
                for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
                    c.drawCentredString(width / 2 + offset[0], top_text_y + 15 - i * 16 + offset[1], line)
        else:
            for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
                c.drawString(width * 0.1 + offset[0], top_text_y + offset[1], title_lines[0])
        
        # 绘制序列号
        for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
            c.drawCentredString(width / 2 + offset[0], serial_number_y + offset[1], current_number)

    def draw_nested_small_box_table(self, c, width, height, theme_text, pieces_per_small_box, 
                                    serial_range, carton_no, remark_text):
        """绘制套盒小箱标表格 - 借鉴分盒模板的表格绘制逻辑"""
        # 表格尺寸和位置 - 上下左右各5mm边距
        table_width = width - 10 * mm
        table_height = height - 10 * mm
        table_x = 5 * mm
        table_y = 5 * mm
        
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
        
        # 绘制Quantity行的分隔线（上层和下层之间）
        quantity_split_y = row_positions[2] + quantity_row_height / 2
        c.line(col_x, quantity_split_y, table_x + table_width, quantity_split_y)
        
        # 表格内容
        font_manager.set_best_font(c, 10, bold=True)
        
        # 计算居中位置
        label_center_x = table_x + label_col_width / 2  # 标签列居中
        data_center_x = col_x + data_col_width / 2      # 数据列居中
        
        # 行1: Item (第5行，从上往下) - 多次绘制加粗
        # 调整文字垂直居中位置 - 减去字体大小的1/3来补偿基线偏移
        font_size = 10
        text_offset = font_size / 3
        item_y = row_positions[4] + base_row_height/2 - text_offset
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], item_y + offset[1], "Item:")
            c.drawCentredString(data_center_x + offset[0], item_y + offset[1], "Paper Cards")
        
        # 行2: Theme (第4行) - 多次绘制加粗
        theme_y = row_positions[3] + base_row_height/2 - text_offset
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], theme_y + offset[1], "Theme:")
        
        # 应用文本清理和换行处理
        clean_theme_text = text_processor.clean_text_for_font(theme_text)
        max_theme_width = data_col_width - 4*mm  # 留出边距
        theme_lines = text_processor.wrap_text_to_fit(c, clean_theme_text, max_theme_width, font_manager.get_chinese_font_name(), 10)
        
        # 绘制主题文本（支持多行） - 多次绘制加粗
        if len(theme_lines) > 1:
            # 多行：调整字体大小并垂直居中
            font_manager.set_best_font(c, 8, bold=True)
            line_height = 10
            # 计算整个文本块的总高度
            total_text_height = (len(theme_lines) - 1) * line_height
            # 重新计算多行文本的垂直居中位置
            # 使用单元格的中心位置，然后调整整个文本块的位置
            cell_center_y = row_positions[3] + base_row_height / 2
            multi_text_offset = 8 / 3  # 8号字体的偏移
            start_y = cell_center_y + total_text_height / 2 - multi_text_offset
            for i, line in enumerate(theme_lines):
                for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                    c.drawCentredString(data_center_x + offset[0], start_y - i * line_height + offset[1], line)
            font_manager.set_best_font(c, 10, bold=True)  # 恢复字体大小
        else:
            for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                c.drawCentredString(data_center_x + offset[0], theme_y + offset[1], theme_lines[0])
        
        # 行3: Quantity (第3行，双倍高度) - 多次绘制加粗
        quantity_label_y = row_positions[2] + quantity_row_height/2 - text_offset
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], quantity_label_y + offset[1], "Quantity:")
        # 上层：票数（在分隔线上方居中）
        upper_y = row_positions[2] + quantity_row_height * 3/4 - text_offset
        pcs_text = f"{pieces_per_small_box}PCS"
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(data_center_x + offset[0], upper_y + offset[1], pcs_text)
        # 下层：序列号范围（在分隔线下方居中）
        lower_y = row_positions[2] + quantity_row_height/4 - text_offset
        clean_serial_range = text_processor.clean_text_for_font(serial_range)
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(data_center_x + offset[0], lower_y + offset[1], clean_serial_range)
        
        # 行4: Carton No (第2行) - 多次绘制加粗
        carton_y = row_positions[1] + base_row_height/2 - text_offset
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], carton_y + offset[1], "Carton No:")
            c.drawCentredString(data_center_x + offset[0], carton_y + offset[1], carton_no)
        
        # 行5: Remark (第1行) - 多次绘制加粗
        remark_y = row_positions[0] + base_row_height/2 - text_offset
        clean_remark_text = text_processor.clean_text_for_font(remark_text)
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], remark_y + offset[1], "Remark:")
            c.drawCentredString(data_center_x + offset[0], remark_y + offset[1], clean_remark_text)

    def draw_nested_large_box_table(self, c, width, height, theme_text, pieces_per_large_box, 
                                    serial_range, carton_no, remark_text):
        """绘制套盒大箱标表格"""
        # 复用小箱标的表格绘制逻辑
        self.draw_nested_small_box_table(c, width, height, theme_text, pieces_per_large_box, 
                                         serial_range, carton_no, remark_text)


# 创建全局实例供nested_box模板使用
nested_box_renderer = NestedBoxRenderer()