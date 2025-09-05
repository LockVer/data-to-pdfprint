"""
分盒盒标模板系统

专门用于生成分盒盒标PDF，采用90x50mm标签格式
分盒盒标特点：
- 上方显示完整主题（如：TAB STREET DRAMA）
- 下方显示编号（如：MOP01002-02），代表第几大箱的第几小箱
- 简洁的双行设计，主题和编号分行显示
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
    from .font_utils import get_chinese_font, get_chinese_bold_font
except ImportError:
    def get_chinese_font():
        return 'Helvetica'
    def get_chinese_bold_font():
        return 'Helvetica'

class DivisionBoxTemplate:
    """分盒盒标模板类 - 90x50mm格式"""
    
    # 标签尺寸 (90x50mm) - 与其他标签保持一致
    LABEL_WIDTH = 90 * mm
    LABEL_HEIGHT = 50 * mm
    
    def __init__(self):
        """初始化模板"""
        self.chinese_font = get_chinese_font()
        self.chinese_bold_font = get_chinese_bold_font()
        
        # 颜色定义
        self.colors = {
            'black': CMYKColor(0, 0, 0, 100),
            'gray': CMYKColor(0, 0, 0, 60),
            'light_gray': CMYKColor(0, 0, 0, 20)
        }
        
    
    def create_division_box_label_data(self, excel_data, box_config):
        """
        根据Excel数据和配置创建分盒盒标数据
        
        Args:
            excel_data: Excel数据字典 {A4, B4, B11, F4}
            box_config: 分盒配置 {
                'min_box_count': 每盒张数,
                'box_per_inner_case': 每小箱盒数,
                'inner_case_per_outer_case': 每大箱小箱数
            }
        
        Returns:
            list: 分盒盒标数据列表
        """
        # 基础数据
        total_sheets = int(excel_data.get('F4', 100))
        min_box_count = box_config.get('min_box_count', 10)
        box_per_inner = box_config.get('box_per_inner_case', 5)
        inner_per_outer = box_config.get('inner_case_per_outer_case', 4)
        
        # 计算总盒数、小箱数、大箱数
        total_boxes = math.ceil(total_sheets / min_box_count)
        total_inner_cases = math.ceil(total_boxes / box_per_inner)
        total_outer_cases = math.ceil(total_inner_cases / inner_per_outer)
        
        print(f"=" * 80)
        print(f"🎯🎯🎯 分盒盒标数据计算 🎯🎯🎯")
        print(f"  总张数: {total_sheets}")
        print(f"  每盒张数: {min_box_count}")
        print(f"  每小箱盒数: {box_per_inner}")
        print(f"  每大箱小箱数: {inner_per_outer}")
        print(f"  总盒数: {total_boxes}")
        print(f"  小箱数: {total_inner_cases}")
        print(f"  大箱数: {total_outer_cases}")
        print(f"=" * 80)
        
        # 基础编号和主题
        base_number = excel_data.get('B11', 'MOP01001')
        # 直接使用"标签名称"关键字右边的数据，不做任何处理
        full_theme = self._search_label_name_data(excel_data)
        
        division_box_labels = []
        
        # 生成每个盒标的数据
        for box_index in range(total_boxes):
            # 计算当前盒属于第几小箱和第几大箱
            inner_case_index = box_index // box_per_inner  # 第几小箱（从0开始）
            outer_case_index = inner_case_index // inner_per_outer  # 第几大箱（从0开始）
            
            # 计算大箱内的小箱序号
            inner_in_outer = (inner_case_index % inner_per_outer) + 1  # 大箱内第几小箱（从1开始）
            
            # 生成分盒编号：基础编号 + 大箱号 + 小箱号
            division_number = self._generate_division_number(
                base_number, outer_case_index, inner_in_outer
            )
            
            # 确保字符串编码正确
            clean_theme = str(full_theme).encode('utf-8').decode('utf-8') if full_theme else 'DEFAULT THEME'
            clean_number = str(division_number).encode('utf-8').decode('utf-8') if division_number else 'MOP01001-01'
            
            label_data = {
                'theme': clean_theme,  # 完整主题
                'number': clean_number,  # 分盒编号
                'box_index': box_index + 1,
                'inner_case_index': inner_case_index + 1,
                'outer_case_index': outer_case_index + 1,
                'inner_in_outer': inner_in_outer
            }
            
            print(f"盒 {box_index + 1}: 第{outer_case_index + 1}大箱第{inner_in_outer}小箱, 编号{division_number}")
            
            division_box_labels.append(label_data)
        
        return division_box_labels
    
    def _search_label_name_data(self, excel_data):
        """
        搜索Excel数据中"标签名称"关键字右边的数据
        直接返回找到的数据，不做任何处理
        """
        print(f"🔍 开始搜索标签名称关键字...")
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
    
    def _draw_multiline_text(self, canvas_obj, text, x, y, max_width, max_height, font_name, font_size, align='left'):
        """
        绘制支持自动换行的多行文本
        
        Args:
            canvas_obj: ReportLab Canvas对象
            text: 要绘制的文本
            x, y: 文本区域左上角坐标
            max_width: 文本区域最大宽度
            max_height: 文本区域最大高度
            font_name: 字体名称
            font_size: 字体大小
            align: 对齐方式 ('left', 'center', 'right')
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
        
        # 计算行高
        line_height = font_size * 1.2
        total_text_height = len(lines) * line_height
        
        # 禁用自动字体调整，保持设定的大字体效果
        # 注释掉自动调整逻辑，确保使用原始大字体
        # if total_text_height > max_height and len(lines) > 1:
        #     adjusted_font_size = max_height / (len(lines) * 1.2)
        #     if adjusted_font_size < font_size:
        #         font_size = max(adjusted_font_size, font_size * 0.6)  # 最小不低于原大小的60%
        #         c.setFont(font_name, font_size)
        #         line_height = font_size * 1.2
        #         total_text_height = len(lines) * line_height
        
        # 计算起始Y坐标（从顶部开始）
        start_y = y - font_size
        
        # 绘制每一行
        for i, line in enumerate(lines):
            line_y = start_y - (i * line_height)
            
            if align == 'center':
                line_width = c.stringWidth(line, font_name, font_size)
                line_x = x + (max_width - line_width) / 2
            elif align == 'right':
                line_width = c.stringWidth(line, font_name, font_size)
                line_x = x + max_width - line_width
            else:  # left
                line_x = x
            
            c.drawString(line_x, line_y, line)
            print(f"绘制文本行 {i+1}: '{line}' 在位置 ({line_x}, {line_y})")
    
    def _draw_bold_multiline_text(self, canvas_obj, text, x, y, max_width, max_height, font_name, font_size, align='left'):
        """
        绘制支持自动换行的粗体多行文本（通过重复绘制实现粗体效果）
        
        Args:
            canvas_obj: ReportLab Canvas对象
            text: 要绘制的文本
            x, y: 文本区域左上角坐标
            max_width: 文本区域最大宽度
            max_height: 文本区域最大高度
            font_name: 字体名称
            font_size: 字体大小
            align: 对齐方式 ('left', 'center', 'right')
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
        
        # 计算行高
        line_height = font_size * 1.2
        total_text_height = len(lines) * line_height
        
        # 禁用自动字体调整，保持设定的大字体效果
        # 注释掉自动调整逻辑，确保使用原始大字体
        # if total_text_height > max_height and len(lines) > 1:
        #     adjusted_font_size = max_height / (len(lines) * 1.2)
        #     if adjusted_font_size < font_size:
        #         font_size = max(adjusted_font_size, font_size * 0.6)  # 最小不低于原大小的60%
        #         c.setFont(font_name, font_size)
        #         line_height = font_size * 1.2
        #         total_text_height = len(lines) * line_height
        
        # 计算起始Y坐标（从顶部开始）
        start_y = y - font_size
        
        # 粗体效果的偏移量 - 多行文本增强3倍
        bold_offsets = [
            (0, 0),      # 原始位置
            (0.9, 0),    # 右偏移，增大3倍
            (0, 0.9),    # 上偏移，增大3倍
            (0.9, 0.9),  # 右上偏移，增大3倍
            (0.45, 0),   # 额外右偏移
            (0, 0.45),   # 额外上偏移
            (0.45, 0.45), # 额外右上偏移
            (0.6, 0.3),  # 更多偏移点
            (0.3, 0.6),  # 更多偏移点
            (0.75, 0.15), # 更多偏移点
            (0.15, 0.75), # 更多偏移点
        ]
        
        # 绘制每一行（多次绘制实现粗体效果）
        for i, line in enumerate(lines):
            line_y = start_y - (i * line_height)
            
            if align == 'center':
                line_width = c.stringWidth(line, font_name, font_size)
                base_line_x = x + (max_width - line_width) / 2
            elif align == 'right':
                line_width = c.stringWidth(line, font_name, font_size)
                base_line_x = x + max_width - line_width
            else:  # left
                base_line_x = x
            
            # 多次绘制实现粗体效果
            for offset_x, offset_y in bold_offsets:
                c.drawString(base_line_x + offset_x, line_y + offset_y, line)
            
            print(f"绘制粗体文本行 {i+1}: '{line}' 在位置 ({base_line_x}, {line_y})")
    
    def _draw_bold_single_line(self, canvas_obj, text, x, y, max_width, font_name, font_size, align='left'):
        """
        绘制单行粗体文本（通过重复绘制实现粗体效果）
        
        Args:
            canvas_obj: ReportLab Canvas对象
            text: 要绘制的文本
            x, y: 文本区域左上角坐标
            max_width: 文本区域最大宽度
            font_name: 字体名称
            font_size: 字体大小
            align: 对齐方式 ('left', 'center', 'right')
        """
        c = canvas_obj
        c.setFont(font_name, font_size)
        
        # 计算文本位置
        text_width = c.stringWidth(text, font_name, font_size)
        
        if align == 'center':
            base_x = x + (max_width - text_width) / 2
        elif align == 'right':
            base_x = x + max_width - text_width
        else:  # left
            base_x = x
        
        # 粗体效果的偏移量 - 单行文本增强3倍
        bold_offsets = [
            (0, 0),      # 原始位置
            (0.9, 0),    # 右偏移，增大3倍
            (0, 0.9),    # 上偏移，增大3倍
            (0.9, 0.9),  # 右上偏移，增大3倍
            (0.45, 0),   # 额外右偏移
            (0, 0.45),   # 额外上偏移
            (0.45, 0.45), # 额外右上偏移
            (0.6, 0.3),  # 更多偏移点
            (0.3, 0.6),  # 更多偏移点
            (0.75, 0.15), # 更多偏移点
            (0.15, 0.75), # 更多偏移点
        ]
        
        # 多次绘制实现粗体效果
        for offset_x, offset_y in bold_offsets:
            c.drawString(base_x + offset_x, y + offset_y, text)
        
        print(f"绘制粗体单行文本: '{text}' 在位置 ({base_x}, {y})")
    
    def _generate_division_number(self, base_number, outer_case_index, inner_in_outer):
        """
        生成分盒编号 (MOP01001-01格式)
        
        Args:
            base_number: 基础编号 (如: MOP01001)
            outer_case_index: 大箱索引 (从0开始)
            inner_in_outer: 大箱内小箱序号 (从1开始)
        
        Returns:
            str: 分盒编号 (如: MOP01001-01)
        """
        try:
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
                # 计算大箱号：基础编号 + 大箱索引
                outer_number = start_num + outer_case_index
                # 保持原数字部分的位数
                width = len(number_part)
                outer_part = f"{prefix_part}{outer_number:0{width}d}"
                # 生成完整的分盒编号：大箱号-小箱号
                result = f"{outer_part}-{inner_in_outer:02d}"
                
                print(f"分盒编号生成: 基础'{base_number}' -> 第{outer_case_index + 1}大箱第{inner_in_outer}小箱 -> {result}")
                return result
            else:
                # 如果无法解析数字，使用简单格式
                outer_part = f"{base_number}_{outer_case_index+1:03d}"
                return f"{outer_part}-{inner_in_outer:02d}"
                
        except Exception as e:
            print(f"分盒编号生成失败: {e}")
            return f"{base_number}_OUTER{outer_case_index+1:03d}-{inner_in_outer:02d}"
    
    def draw_division_box_label_on_canvas(self, canvas_obj, label_data, x, y):
        """
        在Canvas上绘制分盒盒标 - 简洁的双行设计
        
        Args:
            canvas_obj: ReportLab Canvas对象
            label_data: 标签数据字典
            x, y: 标签左下角坐标
        """
        c = canvas_obj
        
        # 重置绘制设置，确保文字正常渲染
        c.setFillColor(black)  # 使用ReportLab标准black颜色
        c.setStrokeColor(black)  # 使用ReportLab标准black颜色
        
        # 标签区域
        label_x = x
        label_y = y
        label_width = self.LABEL_WIDTH
        label_height = self.LABEL_HEIGHT
        
        # 绘制标签边框（可选，用于调试）
        # c.setLineWidth(0.5)
        # c.rect(label_x, label_y, label_width, label_height)
        
        # 主题文字 - 上半部分，居中显示，支持自动换行
        theme = label_data.get('theme', 'DEFAULT THEME')
        theme_text = str(theme) if theme else 'DEFAULT THEME'
        
        # 设置主题字体和大小（使用粗体效果）
        theme_font_size = 18  # 适中的大字体效果，避免重叠
        
        # 绘制主题 - 支持自动换行的多行显示
        # 定义主题文字区域（盒标采用上下分区设计）
        theme_max_width = label_width - 6 * mm  # 左右各留3mm边距
        theme_max_height = label_height * 0.6   # 给主题文字更多空间，60%的高度
        
        # 主题区域的左上角坐标
        theme_x = label_x + 3 * mm  # 左边距
        theme_y = label_y + label_height - 9 * mm  # 从标签顶部向下9mm开始，再往下移动
        
        # 使用多行粗体绘制，支持自动换行和居中对齐
        self._draw_bold_multiline_text(
            c, theme_text, theme_x, theme_y,
            theme_max_width, theme_max_height, 
            self.chinese_font, theme_font_size,
            align='center'  # 居中对齐
        )
        
        # 编号文字 - 下半部分，居中显示，使用粗体效果
        number = label_data.get('number', 'MOP01001-01')
        number_text = str(number) if number else 'MOP01001-01'
        
        # 设置编号字体和大小（使用粗体效果）
        number_font_size = 18  # 适中的大字体效果，与主题保持一致
        
        # 绘制编号（使用粗体效果）
        number_area_width = label_width * 0.9  # 编号区域宽度
        number_area_height = label_height * 0.25  # 编号区域高度
        number_start_x = label_x + (label_width - number_area_width) / 2
        number_start_y = label_y + label_height * 0.25  # 更靠下的位置，增加与主题的间距
        
        self._draw_bold_single_line(
            c, number_text, number_start_x, number_start_y,
            number_area_width, self.chinese_font, number_font_size,
            align='center'
        )
    
    def generate_division_box_labels_pdf(self, excel_data, box_config, output_path):
        """
        生成分盒盒标PDF - 90x50mm页面尺寸
        
        Args:
            excel_data: Excel数据字典
            box_config: 分盒配置
            output_path: 输出路径
            
        Returns:
            dict: 生成结果信息
        """
        # 创建分盒盒标数据
        division_box_data = self.create_division_box_label_data(excel_data, box_config)
        
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
        
        # 输出文件路径 - 分盒盒标命名
        division_box_file = label_folder / f"{clean_customer_name}+{clean_theme}+分盒盒标.pdf"
        
        # 创建PDF - 使用标签本身尺寸作为页面尺寸
        page_size = (self.LABEL_WIDTH, self.LABEL_HEIGHT)  # 90x50mm页面
        c = canvas.Canvas(str(division_box_file), pagesize=page_size)
        
        # 设置PDF元数据
        c.setTitle(f"分盒盒标 - {excel_data.get('B4', 'DEFAULT')}")
        c.setAuthor("数据转PDF打印工具")
        c.setSubject("90x50mm分盒盒标批量打印")
        c.setCreator("分盒盒标生成工具 v1.0")
        c.setKeywords("分盒盒标,标签,PDF,打印")
        
        print(f"分盒盒标页面布局: 90x50mm页面，每页1个标签")
        print(f"总计需要生成 {len(division_box_data)} 个分盒盒标")
        
        # 生成每个分盒盒标
        for i, label_data in enumerate(division_box_data):
            print(f"生成分盒盒标 {i+1}/{len(division_box_data)}: {label_data.get('number', f'LABEL{i+1}')}")
            
            # 在Canvas上直接绘制标签，标签从页面左下角开始
            self.draw_division_box_label_on_canvas(c, label_data, 0, 0)
            
            # 每个标签后都换页（除了最后一个）
            if i < len(division_box_data) - 1:
                c.showPage()
        
        # 保存PDF
        c.save()
        
        print(f"✅ 分盒盒标PDF生成成功: {division_box_file.name}")
        print(f"   总计生成 {len(division_box_data)} 个分盒盒标")
        
        return {
            'division_box_labels': str(division_box_file),
            'folder': str(label_folder),
            'count': len(division_box_data)
        }