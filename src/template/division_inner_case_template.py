"""
分合小箱标模板系统

专门用于生成分合小箱标PDF，采用与盒标相同的格式：A4横向页面，每页一条数据，90x50mm标签居中显示
基于双层循环逻辑：盒张数 * 每小箱盒数(写死1) + 序号的父子编号系统 + 箱号的双层编码系统
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

class DivisionInnerCaseTemplate:
    """分合小箱标模板类 - 与盒标使用相同的90x50mm格式，但采用双层循环逻辑"""
    
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
    
    def create_division_inner_case_label_data(self, excel_data, box_config):
        """
        根据Excel数据和配置创建分合小箱标数据 - 采用双层循环逻辑
        
        Args:
            excel_data: Excel数据字典 {A4, B4, B11, F4}
            box_config: 盒标配置 {min_box_count, box_per_inner_case(固定为1), inner_case_per_outer_case}
        
        Returns:
            list: 分合小箱标数据列表
        """
        # 基础数据
        total_sheets = int(excel_data.get('F4', 100))
        min_box_count = box_config.get('min_box_count', 10)
        box_per_inner = 1  # 分盒模版固定为1
        inner_case_per_outer = box_config.get('inner_case_per_outer_case', 2)
        
        # 计算总盒数和内箱数
        total_boxes = math.ceil(total_sheets / min_box_count)
        total_inner_cases = math.ceil(total_boxes / box_per_inner)
        
        # 分合小箱标的数量计算：盒张数 * 每小箱盒数
        # 盒张数 = min_box_count
        # 每小箱盒数 = box_per_inner（固定为1）
        sheets_per_inner_case = min_box_count * box_per_inner
        
        inner_case_labels = []
        
        for i in range(total_inner_cases):
            # 计算当前内箱的实际张数
            remaining_sheets = total_sheets - (i * sheets_per_inner_case)
            current_sheets = min(sheets_per_inner_case, remaining_sheets)
            
            # 直接搜索"标签名称"关键字右边的数据
            english_theme = self._search_label_name_data(excel_data)
            print(f"搜索到的主题: '{english_theme}'")
            
            # 序号的父子编号系统 - 基于当前内箱的盒数范围
            start_box = i * box_per_inner + 1  # 当前内箱的第一个盒号
            current_boxes_in_case = min(box_per_inner, total_boxes - i * box_per_inner)  # 当前内箱实际包含的盒数
            end_box = start_box + current_boxes_in_case - 1  # 当前内箱的最后一个盒号
            
            # 分合小箱标的编号格式：符合开始号结束号相同的规则
            # 父级编号（包含大箱递增） - 子级编号（小箱递增）
            base_number = excel_data.get('B11', 'DEFAULT001')
            outer_case_index = i // inner_case_per_outer + 1  # 大箱序号（循环递增）
            inner_case_index = i % inner_case_per_outer + 1   # 大箱内的小箱序号（循环递增）
            
            # 生成父级编号：基础编号 + 大箱序号递增
            parent_number = self._generate_parent_number_with_outer_case(base_number, outer_case_index - 1)
            
            # 生成子级编号：父级编号 + 小箱序号
            child_number = f"{parent_number}-{inner_case_index:02d}"
            
            # 开始号和结束号相同（因为每个小箱只有一个编号）
            number_range = f"{child_number}-{child_number}"
            print(f"分合编号生成: 基础'{base_number}' -> 父级'{parent_number}' -> 子级'{child_number}' -> 范围'{number_range}'")
            
            # 箱号格式：简单的大箱-小箱格式（如：1-1）
            carton_no = f"{outer_case_index}-{inner_case_index}"
            print(f"箱号: 大箱{outer_case_index}/小箱{inner_case_index} -> {carton_no}")
            
            # 直接使用搜索到的原始数据，不做编码处理
            label_data = {
                'item': 'Paper Cards',  # 固定值
                'theme': english_theme,  # 直接使用搜索到的主题数据
                'quantity': f"{current_sheets}PCS",  # 小箱张数 = 盒张数 * 每小箱盒数(1)
                'number_range': number_range,  # 父子编号系统的编号范围
                'carton_no': carton_no,  # 双层编码的箱号
                'remark': excel_data.get('A4', '默认客户'),  # 直接使用原始备注数据
                'case_index': i + 1,
                'total_cases': total_inner_cases,
                'outer_case_index': outer_case_index,
                'inner_case_index': inner_case_index
            }
            
            print(f"分合小箱标数据: 第{i+1}个 theme='{english_theme}', range='{number_range}', carton='{carton_no}'")
            
            inner_case_labels.append(label_data)
        
        return inner_case_labels
    
    def _generate_parent_number_with_outer_case(self, base_number, outer_case_index):
        """
        生成包含大箱序号的父级编号
        
        Args:
            base_number: 基础编号 (如: LGM01001)
            outer_case_index: 大箱索引 (从0开始)
        
        Returns:
            str: 父级编号 (如: LGM01002 表示第2个大箱)
        """
        try:
            # 清理基础编号
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
                # 父级编号：基础编号 + 大箱序号递增
                parent_num = start_num + outer_case_index
                # 保持原数字部分的位数
                width = len(number_part)
                result = f"{prefix_part}{parent_num:0{width}d}"
                
                print(f"父级编号生成: 基础'{base_number}' + 大箱索引{outer_case_index} -> '{result}'")
                return result
            else:
                # 如果无法解析数字，使用简单格式
                return f"{base_number}_{outer_case_index+1:03d}"
                
        except Exception as e:
            print(f"父级编号生成失败: {e}")
            return f"{base_number}_{outer_case_index+1:03d}"
    
    def _generate_division_number_by_index(self, base_number, index):
        """
        根据基础编号和索引生成对应的分合编号 - 父子编号系统
        参考常规模版的编号生成逻辑，但针对分合小箱标的特殊需求进行调整
        
        Args:
            base_number: 基础编号 (如: LAN01001)
            index: 索引（从0开始）
        
        Returns:
            str: 生成的分合编号
        """
        try:
            # 提取前缀和数字部分 - 参考常规模版的逻辑
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
                # 分合小箱标的父子编号逻辑：
                # 由于每小箱盒数固定为1，所以编号是连续的
                # 与常规模版保持一致的递增逻辑
                current_number = start_num + index
                # 保持原数字部分的位数，确保格式一致
                width = len(number_part)
                result = f"{prefix_part}{current_number:0{width}d}"
                return result
            else:
                # 如果无法解析数字，使用简单递增，保持与常规模版一致
                return f"{base_number}_{index+1:03d}"
                
        except Exception as e:
            print(f"分合编号生成失败: {e}")
            return f"{base_number}_{index+1:03d}"
    
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
    
    def _create_quantity_cell_content(self, quantity, number_range, full_width=False):
        """
        创建带分隔线的Quantity单元格内容
        
        Args:
            quantity: 数量信息 (如: 2850PCS)
            number_range: 编号范围 (如: LAN01001-LAN01005)
            full_width: 是否跨越整个表格宽度
            
        Returns:
            自定义的单元格内容
        """
        
        class QuantityCell(Flowable):
            def __init__(self, quantity, number_range, full_width=False):
                Flowable.__init__(self)
                self.quantity = quantity
                self.number_range = number_range
                self.width = 95*mm if full_width else 70*mm  # 匹配第二列宽度
                self.height = 16*mm  # 单元格高度
            
            def draw(self):
                canvas = self.canv
                
                # 设置字体
                canvas.setFont('Helvetica-Bold', 10)
                canvas.setFillColor(black)
                
                # 绘制数量（上半部分）
                text_width = canvas.stringWidth(self.quantity, 'Helvetica-Bold', 10)
                x_pos = (self.width - text_width) / 2  # 居中
                canvas.drawString(x_pos, self.height - 5*mm, self.quantity)
                
                # 绘制分隔线 - 跨越整个单元格宽度
                line_y = self.height / 2  # 中间位置
                canvas.setLineWidth(1)
                canvas.line(0, line_y, self.width, line_y)  # 从0到整个宽度
                
                # 绘制编号范围（下半部分）
                range_width = canvas.stringWidth(self.number_range, 'Helvetica-Bold', 10)
                x_pos_range = (self.width - range_width) / 2  # 居中
                canvas.drawString(x_pos_range, 2*mm, self.number_range)
        
        return QuantityCell(quantity, number_range, full_width)
    
    def draw_table_on_canvas(self, canvas_obj, label_data, x, y):
        """
        在Canvas上绘制分合小箱标表格 - 直接占满整个页面
        
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
        
        # 绘制外边框 - 优化线条粗细以精确匹配参考图片
        c.setLineWidth(1.0)  # 外边框线条粗细
        c.rect(table_x, table_y, table_width, table_height)
        
        # 表格行高和列宽定义 - 基于实际表格尺寸
        row_height = table_height / 6  # 6行表格，基于实际表格高度
        col1_width = table_width * 0.3  # 第一列占30%
        col2_width = table_width * 0.7  # 第二列占70%
        col_divider_x = table_x + col1_width  # 列分隔线位置
        
        # 绘制内部表格线条
        c.setLineWidth(0.6)  # 内部线条使用更细的线条，匹配参考图片
        
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
        
        # 设置字体 - 精确匹配参考图片的字体大小
        font_size_label = 9    # 标签列字体，与内容列一致
        font_size_content = 9  # 内容列基础字体，稍微减小 
        font_size_theme = 9    # Theme行字体，与其他内容行一致
        font_size_carton = 9   # Carton No.行字体，保持一致
        
        # 表格内容数据 - 改为6行，Quantity分为两行
        table_rows = [
            ('Item:', label_data.get('item', 'Paper Cards')),
            ('Theme:', label_data.get('theme', 'DEFAULT')),
            ('Quantity:', label_data.get('quantity', '0PCS')),  # Quantity第一行：数量
            ('', label_data.get('number_range', '')),  # Quantity第二行：编号范围，左列空
            ('Carton No.:', label_data.get('carton_no', '1-1')),  # 简单箱号格式
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
                
                # 计算居中位置
                text_width = c.stringWidth(content_text, self.chinese_font, current_size)
                centered_x = content_x + (col2_width - text_width) / 2 - 2 * mm
                
                # 绘制单行粗体文本
                self._draw_bold_text(c, content_text, centered_x, row_y_center - 1 * mm, self.chinese_font, current_size)
                print(f"绘制粗体内容 {i}: '{content_text}' 在位置 ({centered_x}, {row_y_center - 1 * mm})")
    
    def generate_division_inner_case_labels_pdf(self, excel_data, box_config, output_path):
        """
        生成分合小箱标PDF - A4横向页面，每页一条数据，90x50mm标签居中显示
        
        Args:
            excel_data: Excel数据字典
            box_config: 盒标配置
            output_path: 输出路径
            
        Returns:
            dict: 生成结果信息
        """
        # 创建分合小箱标数据
        inner_case_data = self.create_division_inner_case_label_data(excel_data, box_config)
        
        # 创建输出目录
        output_dir = Path(output_path)
        customer_name = excel_data.get('A4', '默认客户')
        theme = excel_data.get('B4', '默认主题')
        
        folder_name = f"{customer_name}+{theme}+标签"
        label_folder = output_dir / folder_name
        label_folder.mkdir(exist_ok=True)
        
        # 输出文件路径 - 分合小箱标命名：客户名称+订单名称+"分合小箱"
        inner_case_file = label_folder / f"{customer_name}+{theme}+分合小箱.pdf"
        
        # 创建PDF - 使用标签本身尺寸作为页面尺寸，与盒标格式保持一致
        page_size = (self.LABEL_WIDTH, self.LABEL_HEIGHT)  # 90x50mm页面
        c = canvas.Canvas(str(inner_case_file), pagesize=page_size)
        
        # 设置PDF/X-3元数据（适用于CMYK打印）
        c.setTitle(f"分合小箱标 - {excel_data.get('B4', 'DEFAULT')}")
        c.setAuthor("数据转PDF打印工具")
        c.setSubject("90x50mm分合小箱标批量打印")
        c.setCreator("分合小箱标生成工具 v1.0")
        c.setKeywords("分合小箱标,标签,PDF/X,CMYK,打印,双层循环")
        
        # PDF/X-3兼容性设置
        try:
            # 设置CMYK颜色空间
            c._doc.catalog.colorSpace = "/DeviceCMYK"
            # 添加PDF/X标识
            c._doc.catalog.GTS_PDFXVersion = "PDF/X-3:2002"
            c._doc.catalog.GTS_PDFXConformance = "PDF/X-3:2002"
        except:
            pass  # 如果ReportLab版本不支有则跳过
        
        print(f"分合小箱标页面布局: 90x50mm页面，每页1个标签")
        print(f"总计需要生成 {len(inner_case_data)} 个分合小箱标")
        
        # 生成每个分合小箱标
        for i, label_data in enumerate(inner_case_data):
            print(f"生成分合小箱标 {i+1}/{len(inner_case_data)}: {label_data.get('carton_no', f'{i+1}-1/1-1')}")
            
            # 在Canvas上直接绘制表格，标签从页面左下角开始
            self.draw_table_on_canvas(c, label_data, 0, 0)
            
            # 每个标签后都换页（除了最后一个）
            if i < len(inner_case_data) - 1:
                c.showPage()
        
        # 保存PDF
        c.save()
        
        print(f"✅ 分合小箱标PDF生成成功: {inner_case_file.name}")
        print(f"   总计生成 {len(inner_case_data)} 个分合小箱标")
        
        return {
            'division_inner_case_labels': str(inner_case_file),
            'folder': str(label_folder),
            'count': len(inner_case_data)
        }