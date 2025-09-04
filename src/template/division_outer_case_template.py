"""
分合大箱标模板系统

专门用于生成分合大箱标PDF，采用与盒标相同的格式：A4横向页面，每页一条数据，90x50mm标签居中显示
基于规则：
- quantity: 大箱内小箱数量 * 盒张数 * 每小箱盒数（写死1）
- serial: 跨范围编号（开始号-结束号）使用父子级编号逻辑
- carton_no: 当前大箱数
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

class DivisionOuterCaseTemplate:
    """分合大箱标模板类 - 与盒标使用相同的90x50mm格式"""
    
    # 标签尺寸 (90x50mm) - 与盒标保持完全一致
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
    
    def create_division_outer_case_label_data(self, excel_data, box_config):
        """
        根据Excel数据和配置创建分合大箱标数据
        
        Args:
            excel_data: Excel数据字典 {A4, B4, B11, F4}
            box_config: 盒标配置 {min_box_count, box_per_inner_case(固定为1), inner_case_per_outer_case}
        
        Returns:
            list: 分合大箱标数据列表
        """
        # 基础数据
        total_sheets = int(excel_data.get('F4', 100))
        min_box_count = box_config.get('min_box_count', 10)
        box_per_inner = 1  # 分盒模版固定为1
        inner_case_per_outer = box_config.get('inner_case_per_outer_case', 2)
        
        # 计算总盒数和内箱数
        total_boxes = math.ceil(total_sheets / min_box_count)
        total_inner_cases = math.ceil(total_boxes / box_per_inner)
        
        # 计算大箱数
        total_outer_cases = math.ceil(total_inner_cases / inner_case_per_outer)
        
        outer_case_labels = []
        
        for i in range(total_outer_cases):
            # 当前大箱数（从1开始）
            current_outer_case = i + 1
            
            # 提取英文主题
            theme_text = excel_data.get('B4', '默认主题')
            english_theme = self._extract_english_theme(theme_text)
            print(f"原始主题: '{theme_text}' -> 提取后: '{english_theme}'")
            
            # 大箱数量计算：大箱内小箱数量 * 盒张数 * 每小箱盒数（写死1）
            # 当前大箱实际包含的小箱数量
            current_inner_cases_in_outer = min(inner_case_per_outer, total_inner_cases - i * inner_case_per_outer)
            outer_case_quantity = current_inner_cases_in_outer * min_box_count * box_per_inner
            
            # 跨范围编号系统：开始号-结束号（父子级编号逻辑）
            base_number = excel_data.get('B11', 'DEFAULT001')
            
            # 开始号：当前大箱数 + 子级编号01（固定）
            start_parent_number = self._generate_parent_number_by_index(base_number, current_outer_case - 1)
            start_serial = f"{start_parent_number}-01"
            
            # 结束号：当前大箱数 + 子级编号为大箱内小箱数量（最大值）
            end_parent_number = start_parent_number  # 同一个大箱的父级编号相同
            end_serial = f"{end_parent_number}-{current_inner_cases_in_outer:02d}"
            
            # 组合跨范围编号
            serial_range = f"{start_serial}-{end_serial}"
            print(f"大箱{current_outer_case}编号生成: 开始'{start_serial}' -> 结束'{end_serial}' -> 范围'{serial_range}'")
            
            # 确保字符串是纯ASCII或正确的UTF-8编码
            clean_theme = str(english_theme).encode('utf-8').decode('utf-8') if english_theme else 'DEFAULT THEME'
            clean_serial = str(serial_range).encode('utf-8').decode('utf-8') if serial_range else 'DEFAULT001'
            clean_remark = str(excel_data.get('A4', '默认客户')).encode('utf-8').decode('utf-8')
            
            label_data = {
                'item': 'Paper Cards',  # 固定值
                'theme': clean_theme,  # 确保编码正确的主题
                'quantity': f"{outer_case_quantity}PCS",  # 大箱数量
                'number_range': clean_serial,  # 跨范围编号
                'carton_no': str(current_outer_case),  # 当前大箱数
                'remark': clean_remark,  # 确保编码正确的备注
                'case_index': current_outer_case,
                'total_cases': total_outer_cases,
                'inner_cases_count': current_inner_cases_in_outer
            }
            
            print(f"分合大箱标数据: 第{current_outer_case}大箱 theme='{clean_theme}', quantity='{outer_case_quantity}PCS', serial='{clean_serial}', carton='{current_outer_case}'")
            
            outer_case_labels.append(label_data)
        
        return outer_case_labels
    
    def _generate_parent_number_by_index(self, base_number, index):
        """
        根据基础编号和大箱索引生成父级编号
        
        Args:
            base_number: 基础编号 (如: LGM01001)
            index: 大箱索引（从0开始）
        
        Returns:
            str: 生成的父级编号
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
                # 父级编号：基础编号 + 大箱索引递增
                current_number = start_num + index
                # 保持原数字部分的位数，确保格式一致
                width = len(number_part)
                result = f"{prefix_part}{current_number:0{width}d}"
                return result
            else:
                # 如果无法解析数字，使用简单递增
                return f"{base_number}_{index+1:03d}"
                
        except Exception as e:
            print(f"父级编号生成失败: {e}")
            return f"{base_number}_{index+1:03d}"
    
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
        在Canvas上绘制分合大箱标表格 - 直接占满整个页面
        
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
        font_size_label = 8    # 标签列字体，稍微减小
        font_size_content = 9  # 内容列基础字体，稍微减小 
        font_size_theme = 9    # Theme行字体，保持一致
        font_size_carton = 9   # Carton No.行字体，保持一致
        
        # 表格内容数据 - 改为6行，Quantity分为两行
        table_rows = [
            ('Item:', label_data.get('item', 'Paper Cards')),
            ('Theme:', label_data.get('theme', 'DEFAULT')),
            ('Quantity:', label_data.get('quantity', '0PCS')),  # Quantity第一行：数量
            ('', label_data.get('number_range', '')),  # Quantity第二行：编号范围，左列空
            ('Carton No.:', label_data.get('carton_no', '1')),  # 当前大箱数
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
    
    def generate_division_outer_case_labels_pdf(self, excel_data, box_config, output_path):
        """
        生成分合大箱标PDF - A4横向页面，每页一条数据，90x50mm标签居中显示
        
        Args:
            excel_data: Excel数据字典
            box_config: 盒标配置
            output_path: 输出路径
            
        Returns:
            dict: 生成结果信息
        """
        # 创建分合大箱标数据
        outer_case_data = self.create_division_outer_case_label_data(excel_data, box_config)
        
        # 创建输出目录
        output_dir = Path(output_path)
        customer_name = excel_data.get('A4', '默认客户')
        theme = excel_data.get('B4', '默认主题')
        
        folder_name = f"{customer_name}+{theme}+标签"
        label_folder = output_dir / folder_name
        label_folder.mkdir(exist_ok=True)
        
        # 输出文件路径 - 分合大箱标命名：客户名称+订单名称+"分合大箱"
        outer_case_file = label_folder / f"{customer_name}+{theme}+分合大箱.pdf"
        
        # 创建PDF - 使用标签本身尺寸作为页面尺寸，与盒标格式保持一致
        page_size = (self.LABEL_WIDTH, self.LABEL_HEIGHT)  # 90x50mm页面
        c = canvas.Canvas(str(outer_case_file), pagesize=page_size)
        
        # 设置PDF/X-3元数据（适用于CMYK打印）
        c.setTitle(f"分合大箱标 - {excel_data.get('B4', 'DEFAULT')}")
        c.setAuthor("数据转PDF打印工具")
        c.setSubject("90x50mm分合大箱标批量打印")
        c.setCreator("分合大箱标生成工具 v1.0")
        c.setKeywords("分合大箱标,标签,PDF/X,CMYK,打印,跨范围编号")
        
        # PDF/X-3兼容性设置
        try:
            # 设置CMYK颜色空间
            c._doc.catalog.colorSpace = "/DeviceCMYK"
            # 添加PDF/X标识
            c._doc.catalog.GTS_PDFXVersion = "PDF/X-3:2002"
            c._doc.catalog.GTS_PDFXConformance = "PDF/X-3:2002"
        except:
            pass  # 如果ReportLab版本不支有则跳过
        
        print(f"分合大箱标页面布局: 90x50mm页面，每页1个标签")
        print(f"总计需要生成 {len(outer_case_data)} 个分合大箱标")
        
        # 生成每个分合大箱标
        for i, label_data in enumerate(outer_case_data):
            print(f"生成分合大箱标 {i+1}/{len(outer_case_data)}: 第{label_data.get('carton_no', '1')}大箱")
            
            # 在Canvas上直接绘制表格，标签从页面左下角开始
            self.draw_table_on_canvas(c, label_data, 0, 0)
            
            # 每个标签后都换页（除了最后一个）
            if i < len(outer_case_data) - 1:
                c.showPage()
        
        # 保存PDF
        c.save()
        
        print(f"✅ 分合大箱标PDF生成成功: {outer_case_file.name}")
        print(f"   总计生成 {len(outer_case_data)} 个分合大箱标")
        
        return {
            'division_outer_case_labels': str(outer_case_file),
            'folder': str(label_folder),
            'count': len(outer_case_data)
        }