"""
套盒小箱标模板系统

专门用于生成套盒小箱标PDF，采用与盒标相同的格式：90x50mm标签
套盒小箱标特点：
- Item: 固定为 "Paper Cards"
- Quantity: 上方显示一套的张数，下方显示这一盒数据的编号
- Carton No.: 显示套号
"""

from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.colors import black, white, CMYKColor
from reportlab.lib.units import mm, cm, inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping
from pathlib import Path
import os
import platform
import math

# 导入统一的字体工具
try:
    from .font_utils import get_chinese_font
except ImportError:
    def get_chinese_font():
        return 'Helvetica' 

class SetBoxInnerCaseTemplate:
    """套盒小箱标模板类 - 90x50mm格式"""
    
    # 标签尺寸 (90x50mm) - 与盒标保持完全一致
    LABEL_WIDTH = 90 * mm
    LABEL_HEIGHT = 50 * mm
    
    def __init__(self):
        """初始化模板"""
        self.chinese_font = get_chinese_font()
        
        # 颜色定义 (CMYK)
        self.colors = {
            'black': CMYKColor(0, 0, 0, 100),
            'gray': CMYKColor(0, 0, 0, 60),
            'light_gray': CMYKColor(0, 0, 0, 20)
        }
    
    def create_set_box_inner_case_label_data(self, excel_data, quantities):
        """
        根据Excel数据和套盒配置创建套盒小箱标数据
        
        Args:
            excel_data: Excel数据字典 {A4, B4, B11, F4}
            quantities: 套盒配置 {
                'min_set_count': 每套张数,
                'boxes_per_set': 几盒为一套,
                'boxes_per_inner_case': 几盒入一小箱,
                'sets_per_outer_case': 几套入一大箱
            }
        
        Returns:
            list: 套盒小箱标数据列表
        """
        # 获取参数，优先使用新参数，兼容旧参数
        total_sheets = int(excel_data.get('F4', 100))
        if 'min_set_count' in quantities:
            min_set_count = quantities['min_set_count']
        elif 'min_box_count' in quantities:
            min_set_count = quantities['min_box_count'] * quantities.get('boxes_per_set', 3)
        else:
            min_set_count = 30
            
        boxes_per_set = quantities.get('boxes_per_set', 3)
        boxes_per_inner_case = quantities.get('boxes_per_inner_case', 6)
        sets_per_outer_case = quantities.get('sets_per_outer_case', 2)
        
        # 计算套数和总盒数
        set_count = math.ceil(total_sheets / min_set_count)
        total_boxes = set_count * boxes_per_set
        
        # 计算小箱数：每个小箱装 boxes_per_inner_case 个盒
        total_inner_cases = math.ceil(total_boxes / boxes_per_inner_case)
        
        print(f"=" * 80)
        print(f"🎯🎯🎯 套盒小箱标数据计算 🎯🎯🎯")
        print(f"  总张数: {total_sheets}")
        print(f"  每套张数: {min_set_count}")
        print(f"  几盒为一套: {boxes_per_set}")
        print(f"  几盒入一小箱: {boxes_per_inner_case}")
        print(f"  套数: {set_count}")
        print(f"  总盒数: {total_boxes}")
        print(f"  小箱数: {total_inner_cases}")
        print(f"=" * 80)
        
        # 基础编号
        base_number = excel_data.get('B11', 'JAW01001')
        
        inner_case_labels = []
        
        # 为每个小箱生成标签
        for inner_case_index in range(total_inner_cases):
            # 计算当前小箱中的盒范围
            start_box_index = inner_case_index * boxes_per_inner_case
            end_box_index = min(start_box_index + boxes_per_inner_case - 1, total_boxes - 1)
            
            # 计算当前小箱包含的盒数
            current_boxes_in_case = end_box_index - start_box_index + 1
            
            # 生成小箱内盒的编号范围
            start_box_number = self._generate_set_box_number(base_number, start_box_index, boxes_per_set)
            end_box_number = self._generate_set_box_number(base_number, end_box_index, boxes_per_set)
            
            # 编号范围
            if start_box_number == end_box_number:
                number_range = start_box_number
            else:
                number_range = f"{start_box_number}-{end_box_number}"
            
            # 计算套号：根据当前小箱包含的盒来确定套号范围
            start_set_index = start_box_index // boxes_per_set
            end_set_index = end_box_index // boxes_per_set
            
            if start_set_index == end_set_index:
                carton_no = f"{start_set_index + 1:02d}"  # 单个套号，格式如 "01"
            else:
                carton_no = f"{start_set_index + 1:02d}-{end_set_index + 1:02d}"  # 套号范围
            
            # 提取主题 - 使用和常规模板完全相同的搜索逻辑
            english_theme = self._search_label_name_data(excel_data)
            
            # 直接使用主题搜索结果，与常规模版保持一致
            clean_theme = str(english_theme) if english_theme else 'JAW'
            clean_range = str(number_range) if number_range else 'JAW01001-01'
            clean_carton_no = str(carton_no) if carton_no else '01'
            clean_remark = str(excel_data.get('A4', '默认客户'))
            
            label_data = {
                'item': 'Paper Cards',  # 固定值
                'theme': clean_theme,  # 确保编码正确的主题
                'quantity': f"{min_set_count}PCS",  # 一套的张数
                'number_range': clean_range,  # 确保编码正确的编号范围
                'carton_no': clean_carton_no,  # 确保编码正确的套号
                'remark': clean_remark,  # 确保编码正确的备注(A4数据)
                'case_index': inner_case_index + 1,
                'total_cases': total_inner_cases
            }
            
            print(f"小箱 {inner_case_index + 1}: 盒{start_box_index + 1}-{end_box_index + 1}, 编号{number_range}, 套号{carton_no}")
            
            inner_case_labels.append(label_data)
        
        return inner_case_labels
    
    def _generate_set_box_number(self, base_number, box_index, boxes_per_set):
        """
        生成套盒编号 (JAW01001-01格式)
        
        Args:
            base_number: 基础编号 (如: JAW01001)
            box_index: 盒索引 (从0开始)
            boxes_per_set: 每套盒数
        
        Returns:
            str: 套盒编号 (如: JAW01001-01)
        """
        try:
            # 计算套索引和套内盒序号
            set_index = box_index // boxes_per_set  # 第几套（从0开始）
            box_in_set = (box_index % boxes_per_set) + 1  # 套内第几盒（从1开始）
            
            # 清理基础编号 - 去掉可能存在的后缀
            clean_base_number = str(base_number).strip()
            if '-' in clean_base_number:
                clean_base_number = clean_base_number.split('-')[0]
            
            # 解析基础编号：前缀 + 数字
            prefix_part = ''
            number_part = ''
            
            for j in range(len(clean_base_number)-1, -1, -1):
                if clean_base_number[j].isdigit():
                    number_part = clean_base_number[j] + number_part
                else:
                    prefix_part = clean_base_number[:j+1]
                    break
            
            if number_part:
                start_num = int(number_part)
                # 计算套号：基础编号 + 套索引
                set_number = start_num + set_index
                # 保持原数字部分的位数
                width = len(number_part)
                set_part = f"{prefix_part}{set_number:0{width}d}"
                # 生成完整的套盒编号：套号-盒在套内序号
                result = f"{set_part}-{box_in_set:02d}"
                
                print(f"套盒编号生成: 盒{box_index + 1} -> 套{set_index + 1}盒{box_in_set} -> {result}")
                return result
            else:
                # 如果无法解析数字，使用简单格式
                set_part = f"{base_number}_{set_index+1:03d}"
                return f"{set_part}-{box_in_set:02d}"
                
        except Exception as e:
            print(f"套盒编号生成失败: {e}")
            return f"{base_number}_SET{set_index+1:03d}-{box_in_set:02d}"
    
    def _search_label_name_data(self, excel_data):
        """
        搜索Excel数据中"标签名称"关键字右边的数据
        直接返回找到的数据，不做任何处理
        """
        print(f"🔍 小箱标开始搜索标签名称关键字...")
        print(f"📋 Excel数据中所有单元格：")
        for key, value in sorted(excel_data.items()):
            if value is not None:
                print(f"   {key}: {repr(value)}")
        
        # 遍历所有Excel数据，查找包含"标签名称"的单元格
        for key, value in excel_data.items():
            if value and "标签名称" in str(value):
                print(f"🔍 在单元格 {key} 找到标签名称关键字: {value}")
                
                # 尝试找到右边单元格的数据
                # 假设key格式为字母+数字，如A4, B5等
                try:
                    import re
                    match = re.match(r'([A-Z]+)(\d+)', key)
                    if match:
                        col_letters = match.group(1)
                        row_number = match.group(2)
                        
                        # 计算右边一列的单元格
                        next_col = self._get_next_column(col_letters)
                        right_cell_key = f"{next_col}{row_number}"
                        
                        print(f"🔍 计算右边单元格: {key} -> {right_cell_key}")
                        
                        # 获取右边单元格的数据
                        right_cell_data = excel_data.get(right_cell_key)
                        if right_cell_data:
                            print(f"✅ 找到标签名称右边数据 ({right_cell_key}): {right_cell_data}")
                            return str(right_cell_data).strip()
                        else:
                            print(f"⚠️  右边单元格 {right_cell_key} 无数据")
                            print(f"📋 检查右边单元格周围的数据：")
                            for check_key, check_value in excel_data.items():
                                if check_key.endswith(row_number) and check_value:
                                    print(f"     {check_key}: {repr(check_value)}")
                except Exception as e:
                    print(f"❌ 解析单元格位置失败: {e}")
        
        # 如果没找到"标签名称"关键字，直接返回B4的数据作为备选
        fallback_theme = excel_data.get('B4', '默认主题')
        print(f"⚠️  未找到标签名称关键字，使用B4备选数据: {fallback_theme}")
        return str(fallback_theme).strip() if fallback_theme else '默认主题'
    
    def _get_next_column(self, col_letters):
        """获取下一列的字母标识"""
        # 将字母转换为数字，加1，再转回字母
        result = 0
        for char in col_letters:
            result = result * 26 + (ord(char) - ord('A') + 1)
        
        result += 1  # 下一列
        
        # 转回字母
        next_col = ''
        while result > 0:
            result -= 1
            next_col = chr(result % 26 + ord('A')) + next_col
            result //= 26
        
        return next_col

    
    def _draw_bold_text(self, canvas_obj, text, x, y, font_name, font_size):
        """
        绘制粗体文本（通过重复绘制实现粗体效果）
        """
        c = canvas_obj
        c.setFont(font_name, font_size)
        
        # 粗体效果的偏移量 - 增加偏移量使字体更粗
        bold_offsets = [
            (0, 0),      # 原始位置
            (0.5, 0),    # 右偏移，增加到0.5
            (0, 0.5),    # 上偏移，增加到0.5  
            (0.5, 0.5),  # 右上偏移
            (0.25, 0),   # 额外的右偏移
            (0, 0.25),   # 额外的上偏移
        ]
        
        # 多次绘制实现粗体效果
        for offset_x, offset_y in bold_offsets:
            c.drawString(x + offset_x, y + offset_y, text)
    
    def _draw_multiline_bold_text(self, canvas_obj, text, x, y, max_width, max_height, font_name, font_size, align='center'):
        """
        绘制支持自动换行的粗体多行文本
        """
        c = canvas_obj
        c.setFont(font_name, font_size)
        
        # 分割文本为单词
        words = text.split()
        lines = []
        current_line = ""
        
        for word in words:
            test_line = current_line + (" " if current_line else "") + word
            test_width = c.stringWidth(test_line, font_name, font_size)
            
            if test_width <= max_width:
                current_line = test_line
            else:
                if current_line:
                    lines.append(current_line)
                    current_line = word
                else:
                    # 单个单词太长，强制换行
                    lines.append(word)
        
        if current_line:
            lines.append(current_line)
        
        # 计算行高 - 为主题文字使用更紧凑的行距
        line_height = font_size * 1.0  # 减小行距，让文字更紧凑
        total_text_height = len(lines) * line_height
        
        # 保持固定字体大小，不做自动调整以确保一致性
        # 如果文本高度超过最大高度，仍保持原字体大小
        # font_size 保持不变，确保所有主题使用相同大小
        
        # 计算起始Y坐标，根据行数决定显示位置
        text_block_height = len(lines) * line_height
        
        if len(lines) == 1:
            # 单行文本：使用与其他行完全相同的Y坐标计算
            # max_height = row_height - 2mm，所以实际行高 = max_height + 2mm
            # 单元格中心应该在 y + (max_height + 2mm) / 2 = y + max_height/2 + 1mm
            cell_center_y = y + max_height / 2 + 1 * mm
            start_y = cell_center_y - 1 * mm  # 与其他行一致的偏移
        else:
            # 多行文本：从顶部开始，留小边距
            start_y = y + max_height - line_height * 0.3
        
        # 粗体效果的偏移量 - 增加偏移量使字体更粗
        bold_offsets = [
            (0, 0),      # 原始位置
            (0.5, 0),    # 右偏移，增加到0.5
            (0, 0.5),    # 上偏移，增加到0.5  
            (0.5, 0.5),  # 右上偏移
            (0.25, 0),   # 额外的右偏移
            (0, 0.25),   # 额外的上偏移
        ]
        
        # 绘制每一行
        for i, line in enumerate(lines):
            line_y = start_y - (i * line_height)
            
            if align == 'center':
                line_width = c.stringWidth(line, font_name, font_size)
                base_x = x + (max_width - line_width) / 2
            elif align == 'right':
                line_width = c.stringWidth(line, font_name, font_size)
                base_x = x + max_width - line_width
            else:  # left
                base_x = x
            
            # 多次绘制实现粗体效果
            for offset_x, offset_y in bold_offsets:
                c.drawString(base_x + offset_x, line_y + offset_y, line)
    
    def draw_set_box_inner_case_table_on_canvas(self, canvas_obj, label_data, x, y):
        """
        在Canvas上绘制套盒小箱标表格
        
        Args:
            canvas_obj: ReportLab Canvas对象
            label_data: 标签数据字典
            x, y: 表格左下角坐标
        """
        c = canvas_obj
        
        # 重置绘制设置
        c.setFillColor(black)
        c.setStrokeColor(black)
        
        # 表格边框 - 留边距
        border_margin = 3 * mm
        table_x = x + border_margin
        table_y = y + border_margin  
        table_width = self.LABEL_WIDTH - 2 * border_margin
        table_height = self.LABEL_HEIGHT - 2 * border_margin
        
        # 绘制外边框
        c.setLineWidth(1.0)
        c.rect(table_x, table_y, table_width, table_height)
        
        # 表格行高和列宽定义
        row_height = table_height / 6  # 6行表格 (Item, Theme, Quantity上, Quantity下, Carton No., Remark)
        col1_width = table_width * 0.3  # 第一列占30%
        col2_width = table_width * 0.7  # 第二列占70%
        col_divider_x = table_x + col1_width
        
        # 绘制内部表格线条
        c.setLineWidth(0.6)
        
        # 水平线 - Quantity行需要特殊处理
        for i in range(1, 6):  # 5条水平线
            line_y = table_y + (i * row_height)
            if i == 3:  # 第3条线（Quantity行内部分隔线），只跨越第二列
                # Quantity行内部的分隔线，只跨越第二列，不跨越第一列（因为第一列是跨行的）
                c.line(col_divider_x, line_y, table_x + table_width, line_y)
            else:  # 其他线条正常画完整横线
                c.line(table_x, line_y, table_x + table_width, line_y)
        
        # 垂直分隔线
        c.line(col_divider_x, table_y, col_divider_x, table_y + table_height)
        
        # 字体设置
        font_size_label = 9    # 标签列字体，与内容列一致
        font_size_content = 9  # 内容列基础字体，稍微减小 
        font_size_theme = 9    # Theme行字体，与其他内容行一致
        font_size_carton = 9   # Carton No.行字体，保持一致
        
        # 表格内容数据
        table_rows = [
            ('Item:', label_data.get('item', 'Paper Cards')),
            ('Theme:', label_data.get('theme', 'JAW')),
            ('Quantity:', label_data.get('quantity', '3780PCS')),  # 套的张数
            ('', label_data.get('number_range', '')),  # 盒编号
            ('Carton No.:', label_data.get('carton_no', '01')),  # 套号
            ('Remark:', label_data.get('remark', ''))
        ]
        
        # 绘制每行内容
        for i, (label, content) in enumerate(table_rows):
            row_y_center = table_y + table_height - (i + 0.5) * row_height
            
            # 第一列 - 标签处理
            if i == 2:  # Quantity行，绘制跨两行的"Quantity:"标签
                c.setFillColor(black)
                c.setFont('Helvetica-Bold', font_size_label)
                label_x = table_x + 2 * mm
                # Quantity标签垂直居中在第3-4行的中间（索引2-3）
                row2_center = table_y + table_height - (2 + 0.5) * row_height  # 第3行中心
                row3_center = table_y + table_height - (3 + 0.5) * row_height  # 第4行中心
                quantity_label_y = (row2_center + row3_center) / 2 - 1 * mm
                # 使用粗体绘制方法，保持与右列一致的粗细
                self._draw_bold_text(c, label, label_x, quantity_label_y, self.chinese_font, font_size_label)
                print(f"绘制跨行 Quantity 标签在位置: {quantity_label_y}")
            elif i == 3:  # 编号行，左列空
                pass
            else:  # 其他行正常绘制左列标签
                if label:
                    c.setFillColor(black)
                    c.setFont('Helvetica-Bold', font_size_label)
                    label_x = table_x + 2 * mm
                    label_y = row_y_center - 1 * mm
                    # 使用粗体绘制方法，保持与右列一致的粗细
                    self._draw_bold_text(c, label, label_x, label_y, self.chinese_font, font_size_label)
            
            # 第二列 - 内容
            content_x = col_divider_x + 2 * mm
            c.setFillColor(black)
            
            # 根据行数设置字体大小并使用粗体绘制
            content_text = str(content) if content else ''
            
            if i == 1:  # Theme行 - 使用多行文本自动换行
                current_size = font_size_theme
                
                # 使用多行粗体文本绘制，支持自动换行
                max_width = col2_width - 4 * mm  # 减去左右边距
                max_height = row_height - 2 * mm  # 减去上下边距
                
                # 绘制多行粗体文本，支持自动换行
                # 传入单元格底部坐标，让多行文本方法内部处理定位
                cell_bottom_y = row_y_center - row_height/2
                self._draw_multiline_bold_text(c, content_text, content_x, cell_bottom_y, 
                                              max_width, max_height, self.chinese_font, current_size, 'center')
                print(f"绘制粗体多行主题: '{content_text}' 字体大小={current_size}pt，自动换行")
                
            else:  # 其他行 - 使用单行粗体文本
                if i == 4:  # Carton No.行
                    current_size = font_size_carton
                else:  # 其他行 (Item, Quantity数量, Quantity编号, Remark)
                    current_size = font_size_content
                
                # 计算居中位置
                text_width = c.stringWidth(content_text, self.chinese_font, current_size)
                centered_x = content_x + (col2_width - text_width) / 2 - 2 * mm
                
                # 绘制单行粗体文本
                self._draw_bold_text(c, content_text, centered_x, row_y_center - 1 * mm, self.chinese_font, current_size)
                print(f"绘制粗体内容 {i}: '{content_text}' 在位置 ({centered_x}, {row_y_center - 1 * mm})")
    
    def generate_set_box_inner_case_labels_pdf(self, excel_data, quantities, output_path):
        """
        生成套盒小箱标PDF
        
        Args:
            excel_data: Excel数据字典
            quantities: 套盒配置
            output_path: 输出路径
            
        Returns:
            dict: 生成结果信息
        """
        # 创建套盒小箱标数据
        inner_case_data = self.create_set_box_inner_case_label_data(excel_data, quantities)
        
        # 创建输出目录
        output_dir = Path(output_path)
        customer_name = excel_data.get('A4', '默认客户')
        theme = excel_data.get('B4', '默认主题')
        
        # 清理文件名中的非法字符
        import re
        clean_customer_name = re.sub(r'[<>:"/\\|?*]', '_', str(customer_name))
        clean_theme = re.sub(r'[<>:"/\\|?*]', '_', str(theme))
        
        folder_name = f"{clean_customer_name}+{clean_theme}+标签"
        label_folder = output_dir / folder_name
        
        try:
            label_folder.mkdir(exist_ok=True)
            print(f"✅ 成功创建输出文件夹: {label_folder}")
        except Exception as e:
            print(f"❌ 创建文件夹失败: {e}")
            raise Exception(f"无法创建输出目录: {e}")
        
        # 输出文件路径 - 套盒小箱标命名
        inner_case_file = label_folder / f"{clean_customer_name}+{clean_theme}+套盒小箱.pdf"
        
        # 创建PDF
        page_size = (self.LABEL_WIDTH, self.LABEL_HEIGHT)
        c = canvas.Canvas(str(inner_case_file), pagesize=page_size)
        
        # 设置PDF元数据
        c.setTitle(f"套盒小箱标 - {excel_data.get('B4', 'DEFAULT')}")
        c.setAuthor("数据转PDF打印工具")
        c.setSubject("90x50mm套盒小箱标批量打印")
        c.setCreator("套盒小箱标生成工具 v1.0")
        c.setKeywords("套盒小箱标,标签,PDF,打印")
        
        print(f"套盒小箱标页面布局: 90x50mm页面，每页1个标签")
        print(f"总计需要生成 {len(inner_case_data)} 个套盒小箱标")
        
        # 生成每个套盒小箱标
        for i, label_data in enumerate(inner_case_data):
            print(f"生成套盒小箱标 {i+1}/{len(inner_case_data)}: 套号{label_data.get('carton_no', f'{i+1}')}")
            
            # 在Canvas上直接绘制表格
            self.draw_set_box_inner_case_table_on_canvas(c, label_data, 0, 0)
            
            # 每个标签后都换页（除了最后一个）
            if i < len(inner_case_data) - 1:
                c.showPage()
        
        # 保存PDF
        c.save()
        
        print(f"✅ 套盒小箱标PDF生成成功: {inner_case_file.name}")
        print(f"   总计生成 {len(inner_case_data)} 个套盒小箱标")
        
        return {
            'set_box_inner_case_labels': str(inner_case_file),
            'folder': str(label_folder),
            'count': len(inner_case_data)
        }