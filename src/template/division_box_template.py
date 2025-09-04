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

class DivisionBoxTemplate:
    """分盒盒标模板类 - 90x50mm格式"""
    
    # 标签尺寸 (90x50mm) - 与其他标签保持一致
    LABEL_WIDTH = 90 * mm
    LABEL_HEIGHT = 50 * mm
    
    def __init__(self):
        """初始化模板"""
        self.chinese_font = self._register_chinese_font()
        
        # 颜色定义
        self.colors = {
            'black': CMYKColor(0, 0, 0, 100),
            'gray': CMYKColor(0, 0, 0, 60),
            'light_gray': CMYKColor(0, 0, 0, 20)
        }
        
    def _register_chinese_font(self):
        """注册中文字体 - 寻找最粗的字体"""
        try:
            system = platform.system()
            
            if system == "Darwin":  # macOS
                # 尝试Helvetica.ttc中的不同字体变体，寻找最粗的
                helvetica_path = "/System/Library/Fonts/Helvetica.ttc"
                if os.path.exists(helvetica_path):
                    print(f"尝试Helvetica.ttc的所有字体变体...")
                    # Helvetica.ttc通常包含多个变体：Regular, Bold, Light等
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
        theme_text = excel_data.get('B4', '默认主题')
        
        # 提取完整主题（不只是英文部分）
        full_theme = self._extract_full_theme(theme_text)
        
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
    
    def _extract_full_theme(self, theme_text):
        """提取完整主题，保持原始格式"""
        if not theme_text:
            return 'DEFAULT THEME'
        
        # 去掉开头的"-"符号，但保留其他格式
        clean_theme = theme_text.lstrip('-').strip()
        
        # 如果主题包含中英文，优先显示英文部分，但保持完整性
        # 对于 "TAB STREET DRAMA" 这样的主题，直接返回
        import re
        
        # 查找英文部分（可能包含空格）
        english_match = re.search(r'[A-Z][A-Z\s\'!-]*[A-Z!]', clean_theme)
        if english_match:
            english_part = english_match.group().strip()
            # 如果英文部分看起来是完整的主题，就使用它
            if len(english_part.split()) >= 2:  # 至少两个单词
                return english_part
        
        # 否则返回清理后的原始主题
        return clean_theme if clean_theme else 'DEFAULT THEME'
    
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
        
        # 主题文字 - 上半部分，居中显示
        theme = label_data.get('theme', 'DEFAULT THEME')
        
        # 设置主题字体和大小
        theme_font_size = 14  # 主题字体大小
        c.setFont(self.chinese_font, theme_font_size)
        
        # 清理字符串编码
        clean_theme = str(theme).encode('latin1', 'replace').decode('latin1')
        
        # 计算主题文字位置 - 垂直居中在上半部分
        theme_text_width = c.stringWidth(clean_theme, self.chinese_font, theme_font_size)
        theme_x = label_x + (label_width - theme_text_width) / 2  # 水平居中
        theme_y = label_y + label_height * 0.65  # 位于标签上半部分
        
        # 绘制主题
        c.drawString(theme_x, theme_y, clean_theme)
        print(f"绘制主题: '{clean_theme}' 在位置 ({theme_x}, {theme_y})")
        
        # 编号文字 - 下半部分，居中显示
        number = label_data.get('number', 'MOP01001-01')
        
        # 设置编号字体和大小
        number_font_size = 12  # 编号字体大小
        c.setFont(self.chinese_font, number_font_size)
        
        # 清理字符串编码
        clean_number = str(number).encode('latin1', 'replace').decode('latin1')
        
        # 计算编号文字位置 - 垂直居中在下半部分
        number_text_width = c.stringWidth(clean_number, self.chinese_font, number_font_size)
        number_x = label_x + (label_width - number_text_width) / 2  # 水平居中
        number_y = label_y + label_height * 0.25  # 位于标签下半部分
        
        # 绘制编号
        c.drawString(number_x, number_y, clean_number)
        print(f"绘制编号: '{clean_number}' 在位置 ({number_x}, {number_y})")
    
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