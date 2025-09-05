"""
分盒模板渲染器
负责分盒模板的所有PDF绘制和渲染逻辑
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import CMYKColor
from reportlab.lib.units import mm

# 导入工具类
from src.utils.font_manager import font_manager
from src.utils.text_processor import text_processor


class SplitBoxRenderer:
    """分盒模板渲染器 - 专门负责分盒模板的PDF绘制逻辑"""
    
    def __init__(self):
        """初始化分盒渲染器"""
        self.renderer_type = "split_box"
    
    def render_appearance_one(self, c, width, top_text, serial_number, top_text_y, serial_number_y):
        """分盒模板盒标外观一渲染"""
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

    def draw_split_box_small_box_table(self, c, width, height, theme_text, pieces_per_small_box, 
                                       serial_range, carton_no, remark_text):
        """绘制分盒小箱标表格"""
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

    def draw_split_box_large_box_table(self, c, width, height, theme_text, pieces_per_box,
                                       small_boxes_per_large_box, serial_range, carton_no, remark_text):
        """绘制分盒大箱标表格"""
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
        # 上层：计算并显示 (张/盒 * 小箱/大箱)PCS
        upper_y = row_positions[2] + quantity_row_height * 3/4
        pcs_count = pieces_per_box * small_boxes_per_large_box  # 张/盒 * 小箱/大箱
        pcs_text = f"{pcs_count}PCS"
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


# 创建全局实例供split_box模板使用
split_box_renderer = SplitBoxRenderer()