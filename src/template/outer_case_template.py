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

class OuterCaseTemplate:
    """外箱标模板类 - 与盒标使用相同的90x50mm格式"""
    
    # 标签尺寸 (90x50mm) - 与盒标和内箱标保持完全一致
    LABEL_WIDTH = 90 * mm
    LABEL_HEIGHT = 50 * mm
    
    def __init__(self):
        """初始化模板"""
        # 使用标准颜色，避免CMYK编码问题
        pass
        
    def _register_chinese_font(self):
        """注册中文字体 - 与盒标模板保持一致"""
        try:
            system = platform.system()
            
            if system == "Darwin":  # macOS
                # 尝试Helvetica.ttc中的不同字体变体，寻找最粗的
                helvetica_path = "/System/Library/Fonts/Helvetica.ttc"
                if os.path.exists(helvetica_path):
                    print(f"尝试Helvetica.ttc的所有字体变体...")
                    # Helvetica.ttc通常包含多个变体：Regular, Bold, Light等
                    # 尝试更多索引，寻找最粗的变体
                    for index in range(20):  # 扩大搜索范围
                        try:
                            font_name = f'HelveticaVariant_{index}'
                            pdfmetrics.registerFont(TTFont(font_name, helvetica_path, subfontIndex=index))
                            print(f"✅ 成功注册Helvetica变体 {index}: {font_name}")
                            # 对于较大的索引值，可能是更粗的变体
                            if index >= 1:  # 通常索引1或更高是Bold变体
                                return font_name
                        except Exception as e:
                            continue
                
                # 备用字体
                other_fonts = [
                    "/System/Library/Fonts/Arial.ttf",
                    "/System/Library/Fonts/STHeiti Medium.ttc"  # 黑体，通常较粗
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
            
            # 提取英文主题
            theme_text = excel_data.get('B4', '默认主题')
            english_theme = self._extract_english_theme(theme_text)
            print(f"原始主题: '{theme_text}' -> 提取后: '{english_theme}'")
            
            # 计算内箱编号范围 - 基于当前外箱包含的内箱范围
            start_inner_case = i * inner_per_outer + 1  # 当前外箱的第一个内箱号
            current_inner_cases_in_outer = min(inner_per_outer, total_inner_cases - i * inner_per_outer)  # 当前外箱实际包含的内箱数
            end_inner_case = start_inner_case + current_inner_cases_in_outer - 1  # 当前外箱的最后一个内箱号
            
            # 生成内箱编号范围 - 基于内箱编号，不是盒编号
            base_number = excel_data.get('B11', 'DEFAULT001')
            start_inner_number = self._generate_inner_case_number(base_number, start_inner_case - 1)  # 开始内箱编号
            end_inner_number = self._generate_inner_case_number(base_number, end_inner_case - 1)     # 结束内箱编号
            inner_case_range = f"{start_inner_number}-{end_inner_number}"  # 内箱编号范围
            print(f"内箱编号生成: 基础'{base_number}' -> 开始内箱'{start_inner_number}' -> 结束内箱'{end_inner_number}' -> 范围'{inner_case_range}'")
            
            # 确保字符串是纯ASCII或正确的UTF-8编码
            clean_theme = str(english_theme).encode('utf-8').decode('utf-8') if english_theme else 'DEFAULT THEME'
            clean_range = str(inner_case_range).encode('utf-8').decode('utf-8') if inner_case_range else 'DEFAULT001-DEFAULT001'
            clean_remark = str(excel_data.get('A4', '默认客户')).encode('utf-8').decode('utf-8')
            
            label_data = {
                'item': 'Paper Cards',  # 固定值
                'theme': clean_theme,  # 确保编码正确的主题
                'quantity': f"{current_sheets}PCS",  # 外箱张数
                'number_range': clean_range,  # 确保编码正确的内箱编号范围
                'carton_no': f"{i+1}/{total_outer_cases}",  # 外箱号：第几个外箱/总外箱数
                'remark': clean_remark,  # 确保编码正确的备注
                'case_index': i + 1,
                'total_cases': total_outer_cases
            }
            
            print(f"清理后的数据: theme='{clean_theme}', inner_case_range='{clean_range}'")
            
            outer_case_labels.append(label_data)
        
        return outer_case_labels
    
    def _generate_inner_case_number(self, base_number, inner_case_index):
        """
        根据基础编号和内箱索引生成对应的内箱编号
        外箱标显示的是内箱的编号范围，不是盒的编号范围
        
        Args:
            base_number: 基础编号 (如: LAN01001)
            inner_case_index: 内箱索引（从0开始）
        
        Returns:
            str: 生成的内箱编号
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
                # 生成内箱编号：基础编号 + 内箱索引
                current_number = start_num + inner_case_index
                # 保持原数字部分的位数
                width = len(number_part)
                result = f"{prefix_part}{current_number:0{width}d}"
                return result
            else:
                # 如果无法解析数字，使用简单递增
                return f"{base_number}_INNER_{inner_case_index+1:03d}"
                
        except Exception as e:
            print(f"内箱编号生成失败: {e}")
            return f"{base_number}_INNER_{inner_case_index+1:03d}"
    
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
        
        # 设置字体 - 与内箱标保持一致的字体大小
        font_size_label = 8    # 标签列字体
        font_size_content = 9  # 内容列基础字体
        font_size_theme = 9    # Theme行字体
        font_size_carton = 9   # Carton No.行字体
        
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
                c.setFont('Helvetica-Bold', font_size_label)
                label_x = table_x + 2 * mm  # 基于表格位置
                # Quantity标签垂直居中在第3-4行的中间
                row3_center = table_y + table_height - (2 + 0.5) * row_height
                row4_center = table_y + table_height - (3 + 0.5) * row_height
                quantity_label_y = (row3_center + row4_center) / 2 - 1 * mm
                c.drawString(label_x, quantity_label_y, label)
                print(f"绘制跨行标签 {i}: '{label}' 在位置 ({label_x}, {quantity_label_y})")
            elif i == 3:  # Quantity第二行，左列空（已在上面绘制）
                pass  # 不绘制左列标签
            else:  # 其他行正常绘制左列标签
                if label:  # 只有当标签非空时才绘制
                    c.setFillColor(black)  # 使用ReportLab的black
                    c.setFont('Helvetica-Bold', font_size_label)
                    label_x = table_x + 2 * mm  # 基于表格位置
                    label_y = row_y_center - 1 * mm
                    c.drawString(label_x, label_y, label)
                    print(f"绘制标签 {i}: '{label}' 在位置 ({label_x}, {label_y})")
            
            # 第二列 - 内容
            content_x = col_divider_x + 2 * mm  # 内容列左边距，增加边距
            c.setFillColor(black)  # 使用ReportLab的black而不是CMYK颜色
            
            # 根据行数设置字体大小
            if i == 1:  # Theme行
                c.setFont('Helvetica-Bold', font_size_theme)
                current_size = font_size_theme
            elif i == 4:  # Carton No.行 (现在是第5行)
                c.setFont('Helvetica-Bold', font_size_carton)  
                current_size = font_size_carton
            else:  # 其他行 (Item, Quantity数量, Quantity编号, Remark)
                c.setFont('Helvetica-Bold', font_size_content)
                current_size = font_size_content
            
            # 清理字符串编码
            clean_content = str(content).encode('latin1', 'replace').decode('latin1') if content else ''
            
            # 用清理后的字符串计算居中位置
            text_width = c.stringWidth(clean_content, 'Helvetica-Bold', current_size)
            centered_x = content_x + (col2_width - text_width) / 2 - 2 * mm
            
            # 不自动换行，保持单行显示
            # 如果文本太长，可以考虑减小字体或截断，但先尝试单行显示
            c.drawString(centered_x, row_y_center - 1 * mm, clean_content)
            print(f"绘制内容 {i}: 原始'{content}' -> 清理后'{clean_content}' 在位置 ({centered_x}, {row_y_center - 1 * mm})")
    
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