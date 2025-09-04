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

class SetBoxInnerCaseTemplate:
    """套盒小箱标模板类 - 90x50mm格式"""
    
    # 标签尺寸 (90x50mm) - 与盒标保持完全一致
    LABEL_WIDTH = 90 * mm
    LABEL_HEIGHT = 50 * mm
    
    def __init__(self):
        """初始化模板"""
        pass
        
    def _register_chinese_font(self):
        """注册中文字体"""
        try:
            system = platform.system()
            
            if system == "Darwin":  # macOS
                helvetica_path = "/System/Library/Fonts/Helvetica.ttc"
                if os.path.exists(helvetica_path):
                    print(f"尝试Helvetica.ttc的所有字体变体...")
                    for index in range(20):
                        try:
                            font_name = f'HelveticaVariant_{index}'
                            pdfmetrics.registerFont(TTFont(font_name, helvetica_path, subfontIndex=index))
                            print(f"✅ 成功注册Helvetica变体 {index}: {font_name}")
                            if index >= 1:  # 通常索引1或更高是Bold变体
                                return font_name
                        except Exception as e:
                            continue
                
                # 备用字体
                other_fonts = [
                    "/System/Library/Fonts/Arial.ttf",
                    "/System/Library/Fonts/STHeiti Medium.ttc"
                ]
                
                for font_path in other_fonts:
                    try:
                        if os.path.exists(font_path):
                            if font_path.endswith('.ttc'):
                                for index in range(5):
                                    try:
                                        font_name = f'ExtraFont_{index}'
                                        pdfmetrics.registerFont(TTFont(font_name, font_path, subfontIndex=index))
                                        print(f"✅ 成功注册额外字体: {font_name}")
                                        return font_name
                                    except:
                                        continue
                            else:
                                font_name = 'ExtraFont'
                                pdfmetrics.registerFont(TTFont(font_name, font_path))
                                print(f"✅ 成功注册额外字体: {font_name}")
                                return font_name
                    except:
                        continue
            
            # 最终备用方案
            print("⚠️ 使用默认Helvetica-Bold字体")
            return 'Helvetica-Bold'
            
        except Exception as e:
            print(f"字体注册失败: {e}")
            return 'Helvetica-Bold'
    
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
            
            # 提取主题
            theme_text = excel_data.get('B4', '默认主题')
            english_theme = self._extract_english_theme(theme_text)
            
            # 确保字符串编码正确 - 与常规内箱标保持一致
            clean_theme = str(english_theme).encode('utf-8').decode('utf-8') if english_theme else 'JAW'
            clean_range = str(number_range).encode('utf-8').decode('utf-8') if number_range else 'JAW01001-01'
            clean_carton_no = str(carton_no).encode('utf-8').decode('utf-8') if carton_no else '01'
            clean_remark = str(excel_data.get('A4', '默认客户')).encode('utf-8').decode('utf-8')
            
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
    
    def _extract_english_theme(self, theme_text):
        """提取英文主题"""
        if not theme_text:
            return 'JAW'
        
        import re
        # 去掉开头的"-"符号
        clean_theme = theme_text.lstrip('-').strip()
        
        # 查找英文部分
        english_patterns = [
            r'[A-Z][A-Z\s\'!]*[A-Z!]',           # 大写字母开头结尾的英文短语
            r'[A-Z]+\'[A-Z\s]+[A-Z!]',           # 带撇号的英文
            r'[A-Z]+[A-Z\s!]*',                  # 任何大写字母组合
            r'[A-Za-z][A-Za-z\s\'!]*[A-Za-z!]'  # 任何英文字母组合
        ]
        
        for pattern in english_patterns:
            match = re.search(pattern, clean_theme)
            if match:
                return match.group().strip()
        
        # 如果找不到英文，返回清理后的主题或默认值
        return clean_theme if clean_theme else 'JAW'
    
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
        font_size_label = 8
        font_size_content = 9
        font_size_theme = 9
        font_size_carton = 9
        
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
                c.drawString(label_x, quantity_label_y, label)
                print(f"绘制跨行 Quantity 标签在位置: {quantity_label_y}")
            elif i == 3:  # 编号行，左列空
                pass
            else:  # 其他行正常绘制左列标签
                if label:
                    c.setFillColor(black)
                    c.setFont('Helvetica-Bold', font_size_label)
                    label_x = table_x + 2 * mm
                    label_y = row_y_center - 1 * mm
                    c.drawString(label_x, label_y, label)
            
            # 第二列 - 内容
            content_x = col_divider_x + 2 * mm
            c.setFillColor(black)
            
            # 根据行数设置字体大小
            if i == 1:  # Theme行
                c.setFont('Helvetica-Bold', font_size_theme)
                current_size = font_size_theme
            elif i == 4:  # Carton No.行
                c.setFont('Helvetica-Bold', font_size_carton)  
                current_size = font_size_carton
            else:  # 其他行
                c.setFont('Helvetica-Bold', font_size_content)
                current_size = font_size_content
            
            # 清理字符串编码
            clean_content = str(content).encode('latin1', 'replace').decode('latin1') if content else ''
            
            # 计算居中位置
            text_width = c.stringWidth(clean_content, 'Helvetica-Bold', current_size)
            centered_x = content_x + (col2_width - text_width) / 2 - 2 * mm
            
            # 绘制文本
            c.drawString(centered_x, row_y_center - 1 * mm, clean_content)
            print(f"绘制套盒小箱标内容 {i}: '{clean_content}' 在位置 ({centered_x}, {row_y_center - 1 * mm})")
    
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