"""
外箱标(大箱标)模板系统

专门用于生成外箱标PDF，采用与盒标相同的格式：90x50mm标签页面，每页一条数据
外箱标包含多个内箱，显示外箱总张数和内箱编号范围
"""

from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.colors import black, white, CMYKColor
from reportlab.lib.units import mm, cm, inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer
from reportlab.platypus.flowables import Flowable
from reportlab.platypus.doctemplate import PageBreak
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

class OuterCaseTemplate:
    """外箱标模板类 - 与盒标使用相同的90x50mm格式"""
    
    # 标签尺寸 (90x50mm) - 与盒标和内箱标保持完全一致
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
    
    def create_outer_case_label_data(self, excel_data, box_config):
        """
        根据Excel数据和配置创建外箱标数据
        
        Args:
            excel_data: Excel数据字典 {A4, B4, B11, F4}
            box_config: 盒标配置 {min_box_count, box_per_inner_case, inner_case_per_outer_case}
        
        Returns:
            list: 外箱标数据列表
        """
        # 基础数据
        total_sheets = int(excel_data.get('F4', 100))
        min_box_count = box_config.get('min_box_count', 10)
        box_per_inner = box_config.get('box_per_inner_case', 5)
        inner_per_outer = box_config.get('inner_case_per_outer_case', 2)
        
        # 计算总盒数和内箱数
        total_boxes = math.ceil(total_sheets / min_box_count)
        total_inner_cases = math.ceil(total_boxes / box_per_inner)
        total_outer_cases = math.ceil(total_inner_cases / inner_per_outer)
        
        # 计算每个外箱的张数 = 分盒张数 * 小箱内的盒数 * 外箱内的内箱数
        sheets_per_outer_case = min_box_count * box_per_inner * inner_per_outer
        
        outer_case_labels = []
        
        for i in range(total_outer_cases):
            # 计算当前外箱的实际张数
            remaining_sheets = total_sheets - (i * sheets_per_outer_case)
            current_sheets = min(sheets_per_outer_case, remaining_sheets)
            
            # 直接使用搜索到的主题，不做任何处理
            theme_text = self._search_label_name_data(excel_data)
            print(f"搜索到的主题: '{theme_text}'")
            
            # 计算盒子编号范围 - 基于当前外箱包含的盒子范围
            # 每个外箱包含的盒子数 = 每个内箱的盒数 * 每个外箱的内箱数
            boxes_per_outer_case = box_per_inner * inner_per_outer
            start_box = i * boxes_per_outer_case + 1  # 当前外箱的第一个盒号
            current_boxes_in_outer = min(boxes_per_outer_case, total_boxes - i * boxes_per_outer_case)  # 当前外箱实际包含的盒数
            end_box = start_box + current_boxes_in_outer - 1  # 当前外箱的最后一个盒号
            
            # 生成盒子编号范围 - 基于盒子编号，不是内箱编号
            base_number = excel_data.get('B11', 'DEFAULT001')
            start_box_number = self._generate_box_number(base_number, start_box - 1)  # 开始盒子编号
            end_box_number = self._generate_box_number(base_number, end_box - 1)     # 结束盒子编号
            box_range = f"{start_box_number}-{end_box_number}"  # 盒子编号范围
            print(f"盒子编号生成: 基础'{base_number}' -> 开始盒子'{start_box_number}' -> 结束盒子'{end_box_number}' -> 范围'{box_range}'")
            
            # 其他字段保持现有的编码处理（仅主题部分直接使用）
            clean_range = str(box_range).encode('utf-8').decode('utf-8') if box_range else 'DEFAULT001-DEFAULT001'
            clean_remark = str(excel_data.get('A4', '默认客户')).encode('utf-8').decode('utf-8')
            
            label_data = {
                'item': 'Paper Cards',  # 固定值
                'theme': theme_text,  # 直接使用搜索到的主题
                'quantity': f"{current_sheets}PCS",  # 外箱张数
                'number_range': clean_range,  # 确保编码正确的盒子编号范围
                'carton_no': f"{i+1}/{total_outer_cases}",  # 外箱号：第几个外箱/总外箱数
                'remark': clean_remark,  # 确保编码正确的备注
                'case_index': i + 1,
                'total_cases': total_outer_cases
            }
            
            print(f"外箱标数据: theme='{theme_text}', box_range='{clean_range}'")
            
            outer_case_labels.append(label_data)
        
        return outer_case_labels
    
    def _generate_box_number(self, base_number, box_index):
        """
        根据基础编号和盒子索引生成对应的盒子编号
        外箱标显示的是盒子的编号范围，不是内箱的编号范围
        
        Args:
            base_number: 基础编号 (如: LAN01001)
            box_index: 盒子索引（从0开始）
        
        Returns:
            str: 生成的盒子编号
        """
        try:
            # 提取前缀和数字部分
            prefix_part = ''
            number_part = ''
            
            # 从后往前找连续的数字
            for j in range(len(base_number)-1, -1, -1):
                if base_number[j].isdigit():
                    number_part = base_number[j] + number_part
                else:
                    prefix_part = base_number[:j+1]
                    break
            
            if number_part:
                start_num = int(number_part)
                # 生成盒子编号：基础编号 + 盒子索引
                current_number = start_num + box_index
                # 保持原数字部分的位数
                width = len(number_part)
                result = f"{prefix_part}{current_number:0{width}d}"
                return result
            else:
                # 如果无法解析数字，使用简单递增
                return f"{base_number}_{box_index+1:03d}"
                
        except Exception as e:
            print(f"盒子编号生成失败: {e}")
            return f"{base_number}_{box_index+1:03d}"
    
    def _search_label_name_data(self, excel_data):
        """
        搜索Excel数据中"标签名称"关键字右边的数据 - 模糊匹配
        直接返回找到的数据，不做任何处理
        """
        print(f"🔍 开始模糊搜索标签名称关键字...")
        print(f"📋 Excel数据中所有单元格：")
        for key, value in sorted(excel_data.items()):
            if value is not None:
                print(f"   {key}: {repr(value)}")
        
        # 先显示所有包含"标签"或"名称"的单元格数据，帮助调试
        print("📋 所有包含'标签'或'名称'的单元格：")
        for key, value in sorted(excel_data.items()):
            try:
                if value is not None and ("标签" in str(value) or "名称" in str(value)):
                    print(f"   {key}: {repr(value)}")
            except Exception as e:
                # 跳过有问题的数据
                continue
        
        # 遍历所有Excel数据，模糊查找包含"标签名称"的单元格
        for key, value in excel_data.items():
            try:
                if value is not None and "标签名称" in str(value):
                    print(f"✅ 在单元格 {key} 找到标签名称关键字: {value}")
                    
                    # 尝试找到右边单元格的数据
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
                            # 不转换为小写，保持原始格式
                            result = str(right_cell_data).strip()
                            print(f"✅ 成功提取标签名称右边数据 ({right_cell_key}): {right_cell_data} -> {result}")
                            return result
                        else:
                            print(f"⚠️  右边单元格 {right_cell_key} 无数据")
                            print(f"📋 检查右边单元格周围的数据：")
                            for check_key, check_value in excel_data.items():
                                if check_key.endswith(row_number) and check_value:
                                    print(f"     {check_key}: {repr(check_value)}")
            except Exception as e:
                # 跳过有问题的数据，继续搜索
                continue
        
        # 如果没找到"标签名称"关键字，使用B4备选数据
        fallback_theme = excel_data.get('B4', '默认主题')
        print(f"⚠️  未找到标签名称关键字，使用B4备选数据: {fallback_theme}")
        return str(fallback_theme).strip() if fallback_theme else '默认主题'
    
    def _get_next_column(self, col_letters):
        """获取下一列的字母标识"""
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
        
        # 粗体效果的偏移量 - 减小偏移量避免重影
        bold_offsets = [
            (0, 0),      # 原始位置
            (0.3, 0),    # 右偏移，减小到0.3
            (0, 0.3),    # 上偏移，减小到0.3  
            (0.3, 0.3),  # 右上偏移
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
        
        # 粗体效果的偏移量 - 减小偏移量避免重影
        bold_offsets = [
            (0, 0),      # 原始位置
            (0.3, 0),    # 右偏移，减小到0.3
            (0, 0.3),    # 上偏移，减小到0.3  
            (0.3, 0.3),  # 右上偏移
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
    
    def _extract_content_after_keyword(self, text, keyword):
        """
        从文本中提取关键词后面的内容
        
        Args:
            text: 包含关键词的文本
            keyword: 要查找的关键词
        
        Returns:
            str: 关键词后面的内容
        """
        if not text or not keyword:
            return text
        
        # 查找关键词位置
        keyword_index = text.find(keyword)
        if keyword_index == -1:
            return text
        
        # 提取关键词后面的内容
        content_after = text[keyword_index + len(keyword):].strip()
        
        # 去掉可能的冒号或其他分隔符
        content_after = content_after.lstrip('：: \t')
        
        return content_after if content_after else text
    
    def _extract_english_theme(self, theme_text):
        """提取英文主题"""
        if not theme_text:
            return 'DEFAULT THEME'
        
        import re
        # 去掉开头的"-"符号
        clean_theme = theme_text.lstrip('-').strip()
        
        # 查找英文部分
        english_patterns = [
            r'[A-Z][A-Z\s\'!]*[A-Z!]',           # 大写字母开头结尾的英文短语
            r'[A-Z]+\'[A-Z\s]+[A-Z!]',           # 带撇号的英文 (如 LADIES NIGHT IN)
            r'[A-Z]+[A-Z\s!]*',                  # 任何大写字母组合
            r'[A-Za-z][A-Za-z\s\'!]*[A-Za-z!]'  # 任何英文字母组合
        ]
        
        for pattern in english_patterns:
            match = re.search(pattern, clean_theme)
            if match:
                return match.group().strip()
        
        return clean_theme if clean_theme else 'DEFAULT THEME'
    
    def draw_table_on_canvas(self, canvas_obj, label_data, x, y):
        """
        在Canvas上绘制外箱标表格 - 与内箱标格式完全一致
        
        Args:
            canvas_obj: ReportLab Canvas对象
            label_data: 标签数据字典
            x, y: 表格左下角坐标
        """
        c = canvas_obj
        
        # 重置绘制设置，确保文字正常渲染 - 使用标准颜色
        c.setFillColor(black)  # 使用ReportLab标准black颜色
        c.setStrokeColor(black)  # 使用ReportLab标准black颜色
        
        # 表格边框 - 加粗外边框，不占满整个页面，留边距
        border_margin = 3 * mm  # 外边距
        table_x = x + border_margin
        table_y = y + border_margin  
        table_width = self.LABEL_WIDTH - 2 * border_margin
        table_height = self.LABEL_HEIGHT - 2 * border_margin
        
        # 绘制外边框 - 与内箱标保持一致的线条粗细
        c.setLineWidth(1.0)  # 外边框线条粗细
        c.rect(table_x, table_y, table_width, table_height)
        
        # 表格行高和列宽定义 - 基于实际表格尺寸
        row_height = table_height / 6  # 6行表格，基于实际表格高度
        col1_width = table_width * 0.3  # 第一列占30%
        col2_width = table_width * 0.7  # 第二列占70%
        col_divider_x = table_x + col1_width  # 列分隔线位置
        
        # 绘制内部表格线条
        c.setLineWidth(0.6)  # 内部线条使用更细的线条，匹配内箱标
        
        # 水平线 - 在Quantity左列跨行区域需要分段绘制
        for i in range(1, 6):  # 5条水平线
            line_y = table_y + (i * row_height)
            if i == 3:  # 第3条线（第3-4行之间），左侧Quantity区域不画线
                # 只画右侧（内容列）的水平线
                c.line(col_divider_x, line_y, table_x + table_width, line_y)
            else:  # 其他线条正常画完整横线
                c.line(table_x, line_y, table_x + table_width, line_y)
        
        # 垂直分隔线 - 完整绘制，因为我们只是左列跨行，右列还是分开的
        c.line(col_divider_x, table_y, col_divider_x, table_y + table_height)
        
        # 设置字体 - 与分合模版保持一致
        font_size_label = 9    # 标签列字体，与分合模版一致
        font_size_content = 9  # 内容列基础字体，与分合模版一致
        font_size_theme = 9    # Theme行字体，与分合模版一致
        font_size_carton = 9   # Carton No.行字体，与分合模版一致
        
        # 表格内容数据 - 改为6行，Quantity分为两行
        table_rows = [
            ('Item:', label_data.get('item', 'Paper Cards')),
            ('Theme:', label_data.get('theme', 'DEFAULT')),
            ('Quantity:', label_data.get('quantity', '0PCS')),  # Quantity第一行：外箱总张数
            ('', label_data.get('number_range', '')),  # Quantity第二行：内箱编号范围，左列空
            ('Carton No.:', label_data.get('carton_no', '1/1')),
            ('Remark:', label_data.get('remark', ''))
        ]
        
        # 绘制每行内容 - 基于实际表格位置
        for i, (label, content) in enumerate(table_rows):
            row_y_center = table_y + table_height - (i + 0.5) * row_height
            
            # 第一列 - 标签处理
            if i == 2:  # Quantity第一行，绘制跨两行的"Quantity:"标签
                c.setFillColor(black)  # 使用ReportLab的black而不是CMYK颜色
                label_x = table_x + 2 * mm  # 基于表格位置
                # Quantity标签垂直居中在第3-4行的中间
                row3_center = table_y + table_height - (2 + 0.5) * row_height
                row4_center = table_y + table_height - (3 + 0.5) * row_height
                quantity_label_y = (row3_center + row4_center) / 2 - 1 * mm
                # 使用粗体绘制方法，保持与右列一致的粗细
                self._draw_bold_text(c, label, label_x, quantity_label_y, self.chinese_font, font_size_label)
                print(f"绘制跨行标签 {i}: '{label}' 在位置 ({label_x}, {quantity_label_y})")
            elif i == 3:  # Quantity第二行，左列空（已在上面绘制）
                pass  # 不绘制左列标签
            else:  # 其他行正常绘制左列标签
                if label:  # 只有当标签非空时才绘制
                    c.setFillColor(black)  # 使用ReportLab的black
                    label_x = table_x + 2 * mm  # 基于表格位置
                    label_y = row_y_center - 1 * mm
                    # 使用粗体绘制方法，保持与右列一致的粗细
                    self._draw_bold_text(c, label, label_x, label_y, self.chinese_font, font_size_label)
                    print(f"绘制标签 {i}: '{label}' 在位置 ({label_x}, {label_y})")
            
            # 第二列 - 内容
            content_x = col_divider_x + 2 * mm  # 内容列左边距，增加边距
            c.setFillColor(black)  # 使用ReportLab的black而不是CMYK颜色
            
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
                
                # 计算居中位置 - 使用微软雅黑字体
                text_width = c.stringWidth(content_text, self.chinese_font, current_size)
                centered_x = content_x + (col2_width - text_width) / 2 - 2 * mm
                
                # 使用粗体绘制方法
                self._draw_bold_text(c, content_text, centered_x, row_y_center - 1 * mm, self.chinese_font, current_size)
                print(f"绘制粗体内容 {i}: '{content_text}' 在位置 ({centered_x}, {row_y_center - 1 * mm})")
    
    def generate_outer_case_labels_pdf(self, excel_data, box_config, output_path):
        """
        生成外箱标PDF - 90x50mm页面，每页一条数据
        
        Args:
            excel_data: Excel数据字典
            box_config: 盒标配置
            output_path: 输出路径
            
        Returns:
            dict: 生成结果信息
        """
        # 创建外箱标数据
        outer_case_data = self.create_outer_case_label_data(excel_data, box_config)
        
        # 创建输出目录
        output_dir = Path(output_path)
        customer_name = excel_data.get('A4', '默认客户')
        theme = excel_data.get('B4', '默认主题')
        
        folder_name = f"{customer_name}+{theme}+标签"
        label_folder = output_dir / folder_name
        label_folder.mkdir(exist_ok=True)
        
        # 输出文件路径 - 大箱标命名：客户名称+订单名称+"大外箱"
        outer_case_file = label_folder / f"{customer_name}+{theme}+大外箱.pdf"
        
        # 创建PDF - 使用标签本身尺寸作为页面尺寸，与盒标格式保持一致
        page_size = (self.LABEL_WIDTH, self.LABEL_HEIGHT)  # 90x50mm页面
        c = canvas.Canvas(str(outer_case_file), pagesize=page_size)
        
        # 设置PDF/X-3元数据（适用于CMYK打印）
        c.setTitle(f"外箱标 - {excel_data.get('B4', 'DEFAULT')}")
        c.setAuthor("数据转PDF打印工具")
        c.setSubject("90x50mm外箱标批量打印")
        c.setCreator("外箱标生成工具 v2.0")
        c.setKeywords("外箱标,标签,PDF/X,CMYK,打印")
        
        # PDF/X-3兼容性设置
        try:
            # 设置CMYK颜色空间
            c._doc.catalog.colorSpace = "/DeviceCMYK"
            # 添加PDF/X标识
            c._doc.catalog.GTS_PDFXVersion = "PDF/X-3:2002"
            c._doc.catalog.GTS_PDFXConformance = "PDF/X-3:2002"
        except:
            pass  # 如果ReportLab版本不支持则跳过
        
        print(f"外箱标页面布局: 90x50mm页面，每页1个标签")
        print(f"总计需要生成 {len(outer_case_data)} 个外箱标")
        
        # 生成每个外箱标
        for i, label_data in enumerate(outer_case_data):
            print(f"生成外箱标 {i+1}/{len(outer_case_data)}: {label_data.get('carton_no', f'{i+1}/1')}")
            
            # 在Canvas上直接绘制表格，标签从页面左下角开始
            self.draw_table_on_canvas(c, label_data, 0, 0)
            
            # 每个标签后都换页（除了最后一个）
            if i < len(outer_case_data) - 1:
                c.showPage()
        
        # 保存PDF
        c.save()
        
        print(f"✅ 外箱标PDF生成成功: {outer_case_file.name}")
        print(f"   总计生成 {len(outer_case_data)} 个外箱标")
        
        return {
            'outer_case_labels': str(outer_case_file),
            'folder': str(label_folder),
            'count': len(outer_case_data)
        }