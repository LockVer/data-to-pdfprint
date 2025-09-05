"""
常规模板渲染器
负责常规模板的所有PDF绘制和渲染逻辑
"""

import sys
import os
from reportlab.pdfgen import canvas
from reportlab.lib.colors import CMYKColor
from reportlab.lib.units import mm

# 导入工具类
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "utils"))
from font_manager import font_manager
from text_processor import text_processor


class RegularRenderer:
    """常规模板渲染器 - 专门负责常规模板的PDF绘制逻辑"""
    
    def __init__(self):
        """初始化常规渲染器"""
        self.renderer_type = "regular"
    
    def render_appearance_one(self, c, width, top_text, serial_number, top_text_y, serial_number_y):
        """渲染外观一：简洁标准样式"""
        # 使用多次绘制实现加粗效果
        c.setFillColor(CMYKColor(0, 0, 0, 1))
        
        # 设置最适合的字体
        font_manager.set_best_font(c, 22, bold=True)
        
        # 清理文本，移除不支持的字符
        clean_top_text = text_processor.clean_text_for_font(top_text)
        
        # 处理上部文本的自动换行和字体大小调整
        top_text_lines = text_processor.wrap_text_to_fit(c, clean_top_text, width - 4*mm, font_manager.get_chinese_font_name(), 22)
        
        # 根据行数调整字体大小和位置
        if len(top_text_lines) > 1:
            # 多行时使用较小字体
            font_manager.set_best_font(c, 18, bold=True)
            line_height = 20  # 行间距
            start_y = top_text_y + (len(top_text_lines) - 1) * line_height / 2
            
            for i, line in enumerate(top_text_lines):
                current_y = start_y - i * line_height
                # 多次绘制增加粗细
                for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
                    c.drawCentredString(width / 2 + offset[0], current_y + offset[1], line)
        else:
            # 单行时使用原字体大小
            for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
                c.drawCentredString(width / 2 + offset[0], top_text_y + offset[1], top_text_lines[0])

        # 重置字体大小绘制序列号
        font_manager.set_best_font(c, 22, bold=True)
        # 下部序列号 - 多次绘制增加粗细  
        for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
            c.drawCentredString(width / 2 + offset[0], serial_number_y + offset[1], serial_number)

    def render_appearance_two(self, c, width, page_size, game_title, ticket_count, serial_number, top_y, bottom_y):
        """渲染外观二：精确的三行布局格式"""
        c.setFillColor(CMYKColor(0, 0, 0, 1))
        font_manager.set_best_font(c, 12, bold=True)
        
        # 获取页面高度
        page_height = page_size[1]
        
        # 左边距 - 统一的左边距
        left_margin = 4 * mm
        
        # 清理文本，移除不支持的字符
        clean_game_title = text_processor.clean_text_for_font(game_title)
        clean_serial_number = text_processor.clean_text_for_font(str(serial_number))
        
        # Game title: 距离顶部更大的距离，支持换行并居中显示
        game_title_y = page_height - 12 * mm  # 距离顶部12mm
        max_title_width = width - 2 * left_margin  # 留出左右边距
        
        # 将Game title文本分为标签部分和内容部分
        title_prefix = "Game title: "
        title_content = clean_game_title
        
        # 计算标签部分宽度
        used_font = font_manager.get_chinese_font_name() or "Helvetica"
        prefix_width = c.stringWidth(title_prefix, used_font, 12)
        
        # 计算内容可用宽度
        content_max_width = max_title_width - prefix_width
        
        # 对内容部分进行换行处理
        title_lines = text_processor.wrap_text_to_fit(c, title_content, content_max_width, font_manager.get_chinese_font_name(), 12)
        
        # 绘制Game title - 加粗效果通过多次绘制实现
        for i, line in enumerate(title_lines):
            current_y = game_title_y - i * 14  # 行间距14点
            if i == 0:
                # 第一行包含"Game title: "前缀，左对齐，多次绘制加粗
                full_line = f"{title_prefix}{line}"
                for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                    c.drawString(left_margin + offset[0], current_y + offset[1], full_line)
            else:
                # 后续换行内容居中显示，多次绘制加粗
                line_width = c.stringWidth(line, used_font, 12)
                center_x = (width - line_width) / 2
                for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                    c.drawString(center_x + offset[0], current_y + offset[1], line)
        
        # Ticket count: 左下区域，增加与Serial的间距，多次绘制加粗
        ticket_count_y = 15 * mm  # 距离底部15mm
        ticket_text = f"Ticket count: {ticket_count}"
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawString(left_margin + offset[0], ticket_count_y + offset[1], ticket_text)
        
        # Serial: 距离底部，多次绘制加粗
        serial_y = 6 * mm  # 距离底部6mm
        serial_text = f"Serial: {clean_serial_number}"
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawString(left_margin + offset[0], serial_y + offset[1], serial_text)

    def draw_small_box_table(self, c, width, height, theme_text, pieces_per_small_box, 
                            serial_range, carton_no, remark_text):
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
        item_y = row_positions[4] + base_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], item_y + offset[1], "Item:")
            c.drawCentredString(data_center_x + offset[0], item_y + offset[1], "Paper Cards")
        
        # 行2: Theme (第4行) - 多次绘制加粗
        theme_y = row_positions[3] + base_row_height/2
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
            # 让文本块在单元格中垂直居中
            start_y = theme_y + total_text_height / 2
            for i, line in enumerate(theme_lines):
                for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                    c.drawCentredString(data_center_x + offset[0], start_y - i * line_height + offset[1], line)
            font_manager.set_best_font(c, 10, bold=True)  # 恢复字体大小
        else:
            for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                c.drawCentredString(data_center_x + offset[0], theme_y + offset[1], theme_lines[0])
        
        # 行3: Quantity (第3行，双倍高度) - 多次绘制加粗
        quantity_label_y = row_positions[2] + quantity_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], quantity_label_y + offset[1], "Quantity:")
        # 上层：票数（在分隔线上方居中）
        upper_y = row_positions[2] + quantity_row_height * 3/4
        pcs_text = f"{pieces_per_small_box}PCS"
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(data_center_x + offset[0], upper_y + offset[1], pcs_text)
        # 下层：序列号范围（在分隔线下方居中）
        lower_y = row_positions[2] + quantity_row_height/4
        clean_serial_range = text_processor.clean_text_for_font(serial_range)
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(data_center_x + offset[0], lower_y + offset[1], clean_serial_range)
        
        # 行4: Carton No (第2行) - 多次绘制加粗
        carton_y = row_positions[1] + base_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], carton_y + offset[1], "Carton No:")
            c.drawCentredString(data_center_x + offset[0], carton_y + offset[1], carton_no)
        
        # 行5: Remark (第1行) - 多次绘制加粗
        remark_y = row_positions[0] + base_row_height/2
        clean_remark_text = text_processor.clean_text_for_font(remark_text)
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], remark_y + offset[1], "Remark:")
            c.drawCentredString(data_center_x + offset[0], remark_y + offset[1], clean_remark_text)

    def render_appearance_two(self, c, width, page_size, game_title, ticket_count, serial_number, top_y, bottom_y):
        """渲染外观二：精确的三行布局格式"""
        c.setFillColor(CMYKColor(0, 0, 0, 1))
        font_manager.set_best_font(c, 12, bold=True)
        
        # 获取页面高度
        page_height = page_size[1]
        
        # 左边距 - 统一的左边距
        left_margin = 4 * mm
        
        # 清理文本，移除不支持的字符
        clean_game_title = text_processor.clean_text_for_font(game_title)
        clean_serial_number = text_processor.clean_text_for_font(str(serial_number))
        
        # Game title: 距离顶部更大的距离，支持换行并居中显示
        game_title_y = page_height - 12 * mm  # 距离顶部12mm
        max_title_width = width - 2 * left_margin  # 留出左右边距
        
        # 将Game title文本分为标签部分和内容部分
        title_prefix = "Game title: "
        title_content = clean_game_title
        
        # 计算标签部分宽度
        used_font = font_manager.get_chinese_font_name() or "Helvetica"
        prefix_width = c.stringWidth(title_prefix, used_font, 12)
        
        # 计算内容可用宽度
        content_max_width = max_title_width - prefix_width
        
        # 对内容部分进行换行处理
        title_lines = text_processor.wrap_text_to_fit(c, title_content, content_max_width, font_manager.get_chinese_font_name(), 12)
        
        # 绘制Game title - 加粗效果通过多次绘制实现
        for i, line in enumerate(title_lines):
            current_y = game_title_y - i * 14  # 行间距14点
            if i == 0:
                # 第一行包含"Game title: "前缀，左对齐，多次绘制加粗
                full_line = f"{title_prefix}{line}"
                for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                    c.drawString(left_margin + offset[0], current_y + offset[1], full_line)
            else:
                # 后续换行内容居中显示，多次绘制加粗
                line_width = c.stringWidth(line, used_font, 12)
                center_x = (width - line_width) / 2
                for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                    c.drawString(center_x + offset[0], current_y + offset[1], line)
        
        # Ticket count: 左下区域，增加与Serial的间距，多次绘制加粗
        ticket_count_y = 15 * mm  # 距离底部15mm
        ticket_text = f"Ticket count: {ticket_count}"
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawString(left_margin + offset[0], ticket_count_y + offset[1], ticket_text)
        
        # Serial: 距离底部，多次绘制加粗
        serial_y = 6 * mm  # 距离底部6mm
        serial_text = f"Serial: {clean_serial_number}"
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawString(left_margin + offset[0], serial_y + offset[1], serial_text)

    def draw_large_box_table(self, c, width, height, theme_text, pieces_per_large_box,
                            serial_range, carton_no, remark_text):
        """绘制大箱标表格"""
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
        item_y = row_positions[4] + base_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], item_y + offset[1], "Item:")
            c.drawCentredString(data_center_x + offset[0], item_y + offset[1], "Paper Cards")
        
        # 行2: Theme (第4行) - 多次绘制加粗
        theme_y = row_positions[3] + base_row_height/2
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
            # 让文本块在单元格中垂直居中
            start_y = theme_y + total_text_height / 2
            for i, line in enumerate(theme_lines):
                for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                    c.drawCentredString(data_center_x + offset[0], start_y - i * line_height + offset[1], line)
            font_manager.set_best_font(c, 10, bold=True)  # 恢复字体大小
        else:
            for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                c.drawCentredString(data_center_x + offset[0], theme_y + offset[1], theme_lines[0])
        
        # 行3: Quantity (第3行，双倍高度) - 多次绘制加粗
        quantity_label_y = row_positions[2] + quantity_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], quantity_label_y + offset[1], "Quantity:")
        # 上层：票数（在分隔线上方居中）
        upper_y = row_positions[2] + quantity_row_height * 3/4
        pcs_text = f"{pieces_per_large_box}PCS"
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(data_center_x + offset[0], upper_y + offset[1], pcs_text)
        # 下层：序列号范围（在分隔线下方居中）
        lower_y = row_positions[2] + quantity_row_height/4
        clean_serial_range = text_processor.clean_text_for_font(serial_range)
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(data_center_x + offset[0], lower_y + offset[1], clean_serial_range)
        
        # 行4: Carton No (第2行) - 多次绘制加粗
        carton_y = row_positions[1] + base_row_height/2
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], carton_y + offset[1], "Carton No:")
            c.drawCentredString(data_center_x + offset[0], carton_y + offset[1], carton_no)
        
        # 行5: Remark (第1行) - 多次绘制加粗
        remark_y = row_positions[0] + base_row_height/2
        clean_remark_text = text_processor.clean_text_for_font(remark_text)
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawCentredString(label_center_x + offset[0], remark_y + offset[1], "Remark:")
            c.drawCentredString(data_center_x + offset[0], remark_y + offset[1], clean_remark_text)


# 创建全局实例供regular模板使用  
regular_renderer = RegularRenderer()