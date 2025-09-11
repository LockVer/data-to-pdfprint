"""
常规模板渲染器
负责常规模板的所有PDF绘制和渲染逻辑
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import CMYKColor
from reportlab.lib.units import mm

# 导入工具类1111111
from src.utils.font_manager import font_manager
from src.utils.text_processor import text_processor


class RegularRenderer:
    """常规模板渲染器 - 专门负责常规模板的PDF绘制逻辑"""
    
    def __init__(self):
        """初始化常规渲染器"""
        self.renderer_type = "regular"
    
    def render_appearance_one(self, c, width, top_text, serial_number, top_text_y, serial_number_y):
        """渲染外观一：简洁标准样式"""
        # 使用粗体字体
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
                c.drawCentredString(width / 2, current_y, line)
        else:
            # 单行时使用原字体大小
            c.drawCentredString(width / 2, top_text_y, top_text_lines[0])

        # 重置字体大小绘制序列号
        font_manager.set_best_font(c, 22, bold=True)
        # 下部序列号
        c.drawCentredString(width / 2, serial_number_y, serial_number)

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
                # 第一行包含"Game title: "前缀，左对齐
                full_line = f"{title_prefix}{line}"
                c.drawString(left_margin, current_y, full_line)
            else:
                # 后续换行内容居中显示
                line_width = c.stringWidth(line, used_font, 12)
                center_x = (width - line_width) / 2
                c.drawString(center_x, current_y, line)
        
        # Ticket count: 左下区域，增加与Serial的间距
        ticket_count_y = 15 * mm  # 距离底部15mm
        ticket_text = f"Ticket count: {ticket_count}"
        c.drawString(left_margin, ticket_count_y, ticket_text)
        
        # Serial: 距离底部
        serial_y = 6 * mm  # 距离底部6mm
        serial_text = f"Serial: {clean_serial_number}"
        c.drawString(left_margin, serial_y, serial_text)

    def draw_small_box_table(self, c, width, height, theme_text, pieces_per_small_box, 
                            serial_range, carton_no, remark_text, template_type="有纸卡备注"):
        """绘制小箱标表格"""
        # 表格尺寸和位置 - 上下左右各5mm边距
        table_width = width - 10 * mm
        table_height = height - 10 * mm
        table_x = 5 * mm
        table_y = 5 * mm
        
        # 根据模版类型调整行数和高度分配
        if template_type == "无纸卡备注":
            # 无纸卡备注：4行 (Item, Quantity, Carton No, Remark)
            total_rows = 4
            base_row_height = table_height / 5  # 分成5份，Quantity占2份
            quantity_row_height = base_row_height * 2  # Quantity行双倍高度
        else:
            # 有纸卡备注：5行 (Item, Theme, Quantity, Carton No, Remark)  
            total_rows = 5
            base_row_height = table_height / 6  # 分成6份，Quantity占2份
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
        
        if template_type == "无纸卡备注":
            # 从底部开始：Remark, Carton No, Quantity(双倍), Item
            for height_val in [base_row_height, base_row_height, quantity_row_height, base_row_height]:
                row_positions.append(current_y)
                current_y += height_val
        else:
            # 从底部开始：Remark, Carton No, Quantity(双倍), Theme, Item
            for height_val in [base_row_height, base_row_height, quantity_row_height, base_row_height, base_row_height]:
                row_positions.append(current_y)
                current_y += height_val
        
        # 绘制行线
        for i in range(1, total_rows):
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
        
        # 调整文字垂直居中位置 - 减去字体大小的1/3来补偿基线偏移
        font_size = 10
        text_offset = font_size / 3
        
        if template_type == "无纸卡备注":
            # 无纸卡备注模版：4行结构
            # 行1: Item (第4行，从上往下) - 显示主题内容
            item_y = row_positions[3] + base_row_height/2 - text_offset
            c.drawCentredString(label_center_x, item_y, "Item:")
            
            # 应用文本清理和换行处理主题内容
            clean_theme_text = text_processor.clean_text_for_font(theme_text)
            max_theme_width = data_col_width - 4*mm  # 留出边距
            theme_lines = text_processor.wrap_text_to_fit(c, clean_theme_text, max_theme_width, font_manager.get_chinese_font_name(), 10)
            
            # 绘制主题文本到Item行（支持多行）
            if len(theme_lines) > 1:
                # 多行：调整字体大小并垂直居中
                font_manager.set_best_font(c, 8, bold=True)
                line_height = 10
                total_text_height = (len(theme_lines) - 1) * line_height
                cell_center_y = row_positions[3] + base_row_height / 2
                multi_text_offset = 8 / 3  # 8号字体的偏移
                start_y = cell_center_y + total_text_height / 2 - multi_text_offset
                for i, line in enumerate(theme_lines):
                    c.drawCentredString(data_center_x, start_y - i * line_height, line)
                font_manager.set_best_font(c, 10, bold=True)  # 恢复字体大小
            else:
                c.drawCentredString(data_center_x, item_y, theme_lines[0])
            
            # 行2: Quantity (第3行，双倍高度)
            quantity_label_y = row_positions[2] + quantity_row_height/2 - text_offset
            c.drawCentredString(label_center_x, quantity_label_y, "Quantity:")
            # 上层：票数
            upper_y = row_positions[2] + quantity_row_height * 3/4 - text_offset
            pcs_text = f"{pieces_per_small_box}PCS"
            c.drawCentredString(data_center_x, upper_y, pcs_text)
            # 下层：序列号范围
            lower_y = row_positions[2] + quantity_row_height/4 - text_offset
            clean_serial_range = text_processor.clean_text_for_font(serial_range)
            c.drawCentredString(data_center_x, lower_y, clean_serial_range)
            
            # 行3: Carton No (第2行)
            carton_y = row_positions[1] + base_row_height/2 - text_offset
            c.drawCentredString(label_center_x, carton_y, "Carton No:")
            c.drawCentredString(data_center_x, carton_y, carton_no)
            
            # 行4: Remark (第1行)
            remark_y = row_positions[0] + base_row_height/2 - text_offset
            clean_remark_text = text_processor.clean_text_for_font(remark_text)
            c.drawCentredString(label_center_x, remark_y, "Remark:")
            c.drawCentredString(data_center_x, remark_y, clean_remark_text)
        
        else:
            # 有纸卡备注模版：5行结构（原有逻辑）
            # 行1: Item (第5行，从上往下) - 显示"Paper Cards"
            item_y = row_positions[4] + base_row_height/2 - text_offset
            c.drawCentredString(label_center_x, item_y, "Item:")
            c.drawCentredString(data_center_x, item_y, "Paper Cards")
            
            # 行2: Theme (第4行) - 显示主题内容
            theme_y = row_positions[3] + base_row_height/2 - text_offset
            c.drawCentredString(label_center_x, theme_y, "Theme:")
            
            # 应用文本清理和换行处理
            clean_theme_text = text_processor.clean_text_for_font(theme_text)
            max_theme_width = data_col_width - 4*mm  # 留出边距
            theme_lines = text_processor.wrap_text_to_fit(c, clean_theme_text, max_theme_width, font_manager.get_chinese_font_name(), 10)
            
            # 绘制主题文本（支持多行）
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
                    c.drawCentredString(data_center_x, start_y - i * line_height, line)
                font_manager.set_best_font(c, 10, bold=True)  # 恢复字体大小
            else:
                c.drawCentredString(data_center_x, theme_y, theme_lines[0])
            
            # 行3: Quantity (第3行，双倍高度)
            quantity_label_y = row_positions[2] + quantity_row_height/2 - text_offset
            c.drawCentredString(label_center_x, quantity_label_y, "Quantity:")
            # 上层：票数（在分隔线上方居中）
            upper_y = row_positions[2] + quantity_row_height * 3/4 - text_offset
            pcs_text = f"{pieces_per_small_box}PCS"
            c.drawCentredString(data_center_x, upper_y, pcs_text)
            # 下层：序列号范围（在分隔线下方居中）
            lower_y = row_positions[2] + quantity_row_height/4 - text_offset
            clean_serial_range = text_processor.clean_text_for_font(serial_range)
            c.drawCentredString(data_center_x, lower_y, clean_serial_range)
            
            # 行4: Carton No (第2行)
            carton_y = row_positions[1] + base_row_height/2 - text_offset
            c.drawCentredString(label_center_x, carton_y, "Carton No:")
            c.drawCentredString(data_center_x, carton_y, carton_no)
            
            # 行5: Remark (第1行)
            remark_y = row_positions[0] + base_row_height/2 - text_offset
            clean_remark_text = text_processor.clean_text_for_font(remark_text)
            c.drawCentredString(label_center_x, remark_y, "Remark:")
            c.drawCentredString(data_center_x, remark_y, clean_remark_text)

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
                # 第一行包含"Game title: "前缀，左对齐
                full_line = f"{title_prefix}{line}"
                c.drawString(left_margin, current_y, full_line)
            else:
                # 后续换行内容居中显示
                line_width = c.stringWidth(line, used_font, 12)
                center_x = (width - line_width) / 2
                c.drawString(center_x, current_y, line)
        
        # Ticket count: 左下区域，增加与Serial的间距
        ticket_count_y = 15 * mm  # 距离底部15mm
        ticket_text = f"Ticket count: {ticket_count}"
        c.drawString(left_margin, ticket_count_y, ticket_text)
        
        # Serial: 距离底部
        serial_y = 6 * mm  # 距离底部6mm
        serial_text = f"Serial: {clean_serial_number}"
        c.drawString(left_margin, serial_y, serial_text)

    def draw_large_box_table(self, c, width, height, theme_text, pieces_per_large_box,
                            serial_range, carton_no, remark_text, template_type="有纸卡备注"):
        """绘制大箱标表格"""
        # 表格尺寸和位置 - 上下左右各5mm边距
        table_width = width - 10 * mm
        table_height = height - 10 * mm
        table_x = 5 * mm
        table_y = 5 * mm
        
        # 根据模版类型调整行数和高度分配
        if template_type == "无纸卡备注":
            # 无纸卡备注：4行 (Item, Quantity, Carton No, Remark)
            total_rows = 4
            base_row_height = table_height / 5  # 分成5份，Quantity占2份
            quantity_row_height = base_row_height * 2  # Quantity行双倍高度
        else:
            # 有纸卡备注：5行 (Item, Theme, Quantity, Carton No, Remark)  
            total_rows = 5
            base_row_height = table_height / 6  # 分成6份，Quantity占2份
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
        
        if template_type == "无纸卡备注":
            # 从底部开始：Remark, Carton No, Quantity(双倍), Item
            for height_val in [base_row_height, base_row_height, quantity_row_height, base_row_height]:
                row_positions.append(current_y)
                current_y += height_val
        else:
            # 从底部开始：Remark, Carton No, Quantity(双倍), Theme, Item
            for height_val in [base_row_height, base_row_height, quantity_row_height, base_row_height, base_row_height]:
                row_positions.append(current_y)
                current_y += height_val
        
        # 绘制行线
        for i in range(1, total_rows):
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
        
        # 调整文字垂直居中位置 - 减去字体大小的1/3来补偿基线偏移
        font_size = 10
        text_offset = font_size / 3
        
        if template_type == "无纸卡备注":
            # 无纸卡备注模版：4行结构
            # 行1: Item (第4行，从上往下) - 显示主题内容
            item_y = row_positions[3] + base_row_height/2 - text_offset
            c.drawCentredString(label_center_x, item_y, "Item:")
            
            # 应用文本清理和换行处理主题内容
            clean_theme_text = text_processor.clean_text_for_font(theme_text)
            max_theme_width = data_col_width - 4*mm  # 留出边距
            theme_lines = text_processor.wrap_text_to_fit(c, clean_theme_text, max_theme_width, font_manager.get_chinese_font_name(), 10)
            
            # 绘制主题文本到Item行（支持多行）
            if len(theme_lines) > 1:
                # 多行：调整字体大小并垂直居中
                font_manager.set_best_font(c, 8, bold=True)
                line_height = 10
                total_text_height = (len(theme_lines) - 1) * line_height
                cell_center_y = row_positions[3] + base_row_height / 2
                multi_text_offset = 8 / 3  # 8号字体的偏移
                start_y = cell_center_y + total_text_height / 2 - multi_text_offset
                for i, line in enumerate(theme_lines):
                    c.drawCentredString(data_center_x, start_y - i * line_height, line)
                font_manager.set_best_font(c, 10, bold=True)  # 恢复字体大小
            else:
                c.drawCentredString(data_center_x, item_y, theme_lines[0])
            
            # 行2: Quantity (第3行，双倍高度)
            quantity_label_y = row_positions[2] + quantity_row_height/2 - text_offset
            c.drawCentredString(label_center_x, quantity_label_y, "Quantity:")
            # 上层：票数
            upper_y = row_positions[2] + quantity_row_height * 3/4 - text_offset
            pcs_text = f"{pieces_per_large_box}PCS"
            c.drawCentredString(data_center_x, upper_y, pcs_text)
            # 下层：序列号范围
            lower_y = row_positions[2] + quantity_row_height/4 - text_offset
            clean_serial_range = text_processor.clean_text_for_font(serial_range)
            c.drawCentredString(data_center_x, lower_y, clean_serial_range)
            
            # 行3: Carton No (第2行)
            carton_y = row_positions[1] + base_row_height/2 - text_offset
            c.drawCentredString(label_center_x, carton_y, "Carton No:")
            c.drawCentredString(data_center_x, carton_y, carton_no)
            
            # 行4: Remark (第1行)
            remark_y = row_positions[0] + base_row_height/2 - text_offset
            clean_remark_text = text_processor.clean_text_for_font(remark_text)
            c.drawCentredString(label_center_x, remark_y, "Remark:")
            c.drawCentredString(data_center_x, remark_y, clean_remark_text)
        
        else:
            # 有纸卡备注模版：5行结构（原有逻辑）
            # 行1: Item (第5行，从上往下) - 显示"Paper Cards"
            item_y = row_positions[4] + base_row_height/2 - text_offset
            c.drawCentredString(label_center_x, item_y, "Item:")
            c.drawCentredString(data_center_x, item_y, "Paper Cards")
            
            # 行2: Theme (第4行) - 显示主题内容
            theme_y = row_positions[3] + base_row_height/2 - text_offset
            c.drawCentredString(label_center_x, theme_y, "Theme:")
            
            # 应用文本清理和换行处理
            clean_theme_text = text_processor.clean_text_for_font(theme_text)
            max_theme_width = data_col_width - 4*mm  # 留出边距
            theme_lines = text_processor.wrap_text_to_fit(c, clean_theme_text, max_theme_width, font_manager.get_chinese_font_name(), 10)
            
            # 绘制主题文本（支持多行）
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
                    c.drawCentredString(data_center_x, start_y - i * line_height, line)
                font_manager.set_best_font(c, 10, bold=True)  # 恢复字体大小
            else:
                c.drawCentredString(data_center_x, theme_y, theme_lines[0])
            
            # 行3: Quantity (第3行，双倍高度)
            quantity_label_y = row_positions[2] + quantity_row_height/2 - text_offset
            c.drawCentredString(label_center_x, quantity_label_y, "Quantity:")
            # 上层：票数（在分隔线上方居中）
            upper_y = row_positions[2] + quantity_row_height * 3/4 - text_offset
            pcs_text = f"{pieces_per_large_box}PCS"
            c.drawCentredString(data_center_x, upper_y, pcs_text)
            # 下层：序列号范围（在分隔线下方居中）
            lower_y = row_positions[2] + quantity_row_height/4 - text_offset
            clean_serial_range = text_processor.clean_text_for_font(serial_range)
            c.drawCentredString(data_center_x, lower_y, clean_serial_range)
            
            # 行4: Carton No (第2行)
            carton_y = row_positions[1] + base_row_height/2 - text_offset
            c.drawCentredString(label_center_x, carton_y, "Carton No:")
            c.drawCentredString(data_center_x, carton_y, carton_no)
            
            # 行5: Remark (第1行)
            remark_y = row_positions[0] + base_row_height/2 - text_offset
            clean_remark_text = text_processor.clean_text_for_font(remark_text)
            c.drawCentredString(label_center_x, remark_y, "Remark:")
            c.drawCentredString(data_center_x, remark_y, clean_remark_text)

    def render_empty_box_label(self, c, width, height, chinese_name, remark_text):
        """渲染空箱标签 - 用于小箱标和大箱标的第一页（有纸卡备注）"""
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
        
        # 文本偏移量
        text_offset = 3
        
        # 行1: Item (第5行，从上往下)
        item_y = row_positions[4] + base_row_height/2 - text_offset
        c.drawCentredString(label_center_x, item_y, "Item:")
        c.drawCentredString(data_center_x, item_y, "Paper Cards")
        
        # 行2: Theme (第4行) - 使用用户输入的中文名称
        theme_y = row_positions[3] + base_row_height/2 - text_offset
        clean_chinese_name = text_processor.clean_text_for_font(chinese_name)
        c.drawCentredString(label_center_x, theme_y, "Theme:")
        c.drawCentredString(data_center_x, theme_y, clean_chinese_name)
        
        # 行3: Quantity (第3行，双倍高度) - 保持空白
        quantity_label_y = row_positions[2] + quantity_row_height/2 - text_offset
        c.drawCentredString(label_center_x, quantity_label_y, "Quantity:")
            # 数据列保持空白
        
        # 行4: Carton No (第2行) - 保持空白
        carton_y = row_positions[1] + base_row_height/2 - text_offset
        c.drawCentredString(label_center_x, carton_y, "Carton No:")
            # 数据列保持空白
        
        # 行5: Remark (第1行)
        remark_y = row_positions[0] + base_row_height/2 - text_offset
        clean_remark_text = text_processor.clean_text_for_font(remark_text)
        c.drawCentredString(label_center_x, remark_y, "Remark:")
        c.drawCentredString(data_center_x, remark_y, clean_remark_text)

    def render_empty_box_label_no_paper_card(self, c, width, height, chinese_name, remark_text):
        """渲染空箱标签 - 用于小箱标和大箱标的第一页（无纸卡备注）"""
        # 表格尺寸和位置 - 上下左右各5mm边距
        table_width = width - 10 * mm
        table_height = height - 10 * mm
        table_x = 5 * mm
        table_y = 5 * mm
        
        # 无纸卡备注：4行，Quantity行占2/5，其他3行各占1/5
        base_row_height = table_height / 5
        quantity_row_height = base_row_height * 2  # Quantity行双倍高度
        
        # 列宽 (标签列:数据列 = 1:2)
        label_col_width = table_width / 3
        data_col_width = table_width * 2 / 3
        
        # 绘制表格边框
        c.setStrokeColor(CMYKColor(0, 0, 0, 1))
        c.setLineWidth(1)
        c.rect(table_x, table_y, table_width, table_height)
        
        # 计算各行的Y坐标 - 从底部开始：Remark, Carton No, Quantity(双倍), Item
        row_positions = []
        current_y = table_y
        for height_val in [base_row_height, base_row_height, quantity_row_height, base_row_height]:
            row_positions.append(current_y)
            current_y += height_val
        
        # 绘制行线
        for i in range(1, 4):
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
        
        # 文本偏移量
        text_offset = 3
        
        # 行1: Item (第4行，从上往下) - 显示中文名称
        item_y = row_positions[3] + base_row_height/2 - text_offset
        c.drawCentredString(label_center_x, item_y, "Item:")
        
        # 应用文本清理和换行处理
        clean_chinese_name = text_processor.clean_text_for_font(chinese_name)
        max_theme_width = data_col_width - 4*mm  # 留出边距
        theme_lines = text_processor.wrap_text_to_fit(c, clean_chinese_name, max_theme_width, font_manager.get_chinese_font_name(), 10)
        
        # 绘制主题文本（支持多行） - 多次绘制加粗
        if len(theme_lines) > 1:
            # 多行：调整字体大小并垂直居中
            font_manager.set_best_font(c, 8, bold=True)
            line_height = 10
            # 计算整个文本块的总高度
            total_text_height = (len(theme_lines) - 1) * line_height
            # 重新计算多行文本的垂直居中位置
            cell_center_y = row_positions[3] + base_row_height / 2
            multi_text_offset = 8 / 3  # 8号字体的偏移
            start_y = cell_center_y + total_text_height / 2 - multi_text_offset
            for i, line in enumerate(theme_lines):
                c.drawCentredString(data_center_x, start_y - i * line_height, line)
            font_manager.set_best_font(c, 10, bold=True)  # 恢复字体大小
        else:
            c.drawCentredString(data_center_x, item_y, theme_lines[0])
        
        # 行2: Quantity (第3行，双倍高度) - 保持空白
        quantity_label_y = row_positions[2] + quantity_row_height/2 - text_offset
        c.drawCentredString(label_center_x, quantity_label_y, "Quantity:")
            # 数据列保持空白
        
        # 行3: Carton No (第2行) - 保持空白
        carton_y = row_positions[1] + base_row_height/2 - text_offset
        c.drawCentredString(label_center_x, carton_y, "Carton No:")
            # 数据列保持空白
        
        # 行4: Remark (第1行)
        remark_y = row_positions[0] + base_row_height/2 - text_offset
        clean_remark_text = text_processor.clean_text_for_font(remark_text)
        c.drawCentredString(label_center_x, remark_y, "Remark:")
        c.drawCentredString(data_center_x, remark_y, clean_remark_text)

    def render_blank_first_page(self, c, width, height, chinese_name):
        """渲染常规模版盒标的空白首页 - 仅显示中文标题"""
        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)
        
        # 清理中文文本
        clean_chinese_name = text_processor.clean_text_for_font(chinese_name)
        
        # 设置大字体用于首页标题
        font_size = 35
        font_manager.set_best_font(c, font_size, bold=True)
        
        # 计算页面中央位置
        center_x = width / 2
        center_y = height / 2
        
        # 计算最大文本宽度（页面宽度的80%）
        max_width = width * 0.8
        
        # 检查文本宽度并进行中文字符级换行
        current_font_name = font_manager.get_chinese_font_name()
        text_width = c.stringWidth(clean_chinese_name, current_font_name, font_size)
        
        if text_width > max_width:
            # 需要换行：使用字符级别分割（适用于中文）
            title_lines = self._wrap_chinese_text_by_chars(c, clean_chinese_name, max_width, current_font_name, font_size)
            
            if len(title_lines) > 1:
                # 多行：调整字体大小
                smaller_font_size = 25
                font_manager.set_best_font(c, smaller_font_size, bold=True)
                # 重新计算换行（基于新的字体大小）
                title_lines = self._wrap_chinese_text_by_chars(c, clean_chinese_name, max_width, current_font_name, smaller_font_size)
                
                # 计算行高和总高度
                line_height = smaller_font_size * 1.2  # 行高为字体大小的1.2倍
                total_height = (len(title_lines) - 1) * line_height
                
                # 计算起始Y位置（垂直居中）
                start_y = center_y + total_height / 2
                
                # 绘制每一行，居中显示 - 多次绘制实现加粗效果
                for i, line in enumerate(title_lines):
                    line_y = start_y - i * line_height
                    for offset in [(-0.5, 0), (0.5, 0), (0, -0.5), (0, 0.5), (0, 0)]:
                        c.drawCentredString(center_x + offset[0], line_y + offset[1], line)
            else:
                # 单行但需要小字体
                for offset in [(-0.5, 0), (0.5, 0), (0, -0.5), (0, 0.5), (0, 0)]:
                    c.drawCentredString(center_x + offset[0], center_y + offset[1], title_lines[0])
        else:
            # 单行：使用原始大字体，居中显示 - 多次绘制实现加粗效果
            for offset in [(-0.5, 0), (0.5, 0), (0, -0.5), (0, 0.5), (0, 0)]:
                c.drawCentredString(center_x + offset[0], center_y + offset[1], clean_chinese_name)

    def _wrap_chinese_text_by_chars(self, c, text, max_width, font_name, font_size):
        """按字符级别换行中文文本（适用于没有空格分隔的中文）"""
        if not text:
            return [""]
        
        lines = []
        current_line = ""
        
        for char in text:
            test_line = current_line + char
            text_width = c.stringWidth(test_line, font_name, font_size)
            
            if text_width <= max_width:
                current_line = test_line
            else:
                if current_line:
                    lines.append(current_line)
                    current_line = char
                else:
                    # 如果单个字符都超宽，强制添加
                    lines.append(char)
                    current_line = ""
        
        if current_line:
            lines.append(current_line)
        
        return lines if lines else [text]

    def render_blank_first_page_appearance_two(self, c, width, height, chinese_name):
        """渲染常规模版外观2的空白首页 - 完全按照外观2格式"""
        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)
        
        # 设置字体 - 与外观2保持一致
        font_manager.set_best_font(c, 12, bold=True)
        
        # 左边距 - 与外观2保持一致
        left_margin = 4 * mm
        
        # 清理中文文本
        clean_chinese_name = text_processor.clean_text_for_font(chinese_name)
        
        # Game title: 距离顶部12mm
        game_title_y = height - 12 * mm
        max_title_width = width - 2 * left_margin  # 留出左右边距
        
        # 将Game title文本分为标签部分和内容部分
        title_prefix = "Game title: "
        title_content = clean_chinese_name
        
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
                # 后续行只包含内容，需要考虑前缀宽度的缩进
                indent_x = left_margin + prefix_width
                for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
                    c.drawString(indent_x + offset[0], current_y + offset[1], line)
        
        # Ticket count: 空行 - 使用与正常外观2完全相同的位置
        ticket_count_y = 15 * mm  # 距离底部15mm，与正常外观2一致
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawString(left_margin + offset[0], ticket_count_y + offset[1], "Ticket count:")
        
        # Serial: 空行 - 使用与正常外观2完全相同的位置
        serial_y = 6 * mm  # 距离底部6mm，与正常外观2一致
        for offset in [(-0.2, 0), (0.2, 0), (0, -0.2), (0, 0.2), (0, 0)]:
            c.drawString(left_margin + offset[0], serial_y + offset[1], "Serial:")


# 创建全局实例供regular模板使用  
regular_renderer = RegularRenderer()