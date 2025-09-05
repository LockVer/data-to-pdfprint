"""
盒标/箱标模板系统

专门用于生成不同规格的盒标和箱标PDF，支持德克斯助手类型的标签格式
"""

from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.colors import black, white, CMYKColor
from reportlab.lib.units import mm, cm, inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping
from reportlab.graphics.barcode import createBarcodeDrawing
from pathlib import Path
import os
import platform
import math

# 内箱标和外箱标现在由GUI统一调用，不在此处生成

class BoxLabelTemplate:
    """盒标/箱标模板类"""
    
    # 标签尺寸 (90x50mm)
    LABEL_WIDTH = 90 * mm
    LABEL_HEIGHT = 50 * mm
    
    def __init__(self):
        """初始化模板"""
        self.chinese_font = self._register_chinese_font()
        
        # 颜色定义 (CMYK)
        self.colors = {
            'black': CMYKColor(0, 0, 0, 100),
            'gray': CMYKColor(0, 0, 0, 60),
            'light_gray': CMYKColor(0, 0, 0, 20)
        }
        
        # 内箱标模板由GUI统一管理
        
    def _register_chinese_font(self):
        """注册中文字体 - 优先微软雅黑，避免中文标点乱码"""
        try:
            system = platform.system()
            
            if system == "Darwin":  # macOS
                # 中文字体路径（按优先级排序）
                chinese_fonts = [
                    ("/System/Library/Fonts/Microsoft YaHei.ttf", "MicrosoftYaHei"),  # 微软雅黑（推荐）
                    ("/System/Library/Fonts/Supplemental/Microsoft YaHei.ttf", "MicrosoftYaHei"), # 备用路径
                    ("/System/Library/Fonts/STHeiti Medium.ttc", "STHeiti"),  # 华文黑体
                    ("/System/Library/Fonts/STHeiti Light.ttc", "STHeitiLight"), # 华文黑体细体
                    ("/System/Library/Fonts/PingFang.ttc", "PingFang"),  # 苹方
                    ("/System/Library/Fonts/Arial Unicode MS.ttf", "ArialUnicode")  # Arial Unicode（支持中文）
                ]
                
                for font_path, font_base_name in chinese_fonts:
                    if os.path.exists(font_path):
                        print(f"🔍 尝试中文字体: {font_path}")
                        try:
                            if font_path.endswith('.ttc'):
                                # TTC文件尝试多个索引
                                for index in range(10):
                                    try:
                                        font_name = f'{font_base_name}_{index}'
                                        pdfmetrics.registerFont(TTFont(font_name, font_path, subfontIndex=index))
                                        print(f"✅ 成功注册中文字体: {font_name}")
                                        return font_name
                                    except Exception as e:
                                        continue
                            else:
                                # TTF文件直接注册
                                font_name = font_base_name
                                pdfmetrics.registerFont(TTFont(font_name, font_path))
                                print(f"✅ 成功注册中文字体: {font_name}")
                                return font_name
                        except Exception as e:
                            print(f"字体注册失败 {font_path}: {e}")
                            continue
            
            elif system == "Windows":  # Windows系统
                windows_fonts = [
                    ("C:/Windows/Fonts/msyh.ttf", "MicrosoftYaHei"),  # 微软雅黑
                    ("C:/Windows/Fonts/msyhbd.ttf", "MicrosoftYaHeiBold"),  # 微软雅黑粗体
                    ("C:/Windows/Fonts/simsun.ttc", "SimSun"),  # 宋体
                    ("C:/Windows/Fonts/simhei.ttf", "SimHei")  # 黑体
                ]
                
                for font_path, font_name in windows_fonts:
                    if os.path.exists(font_path):
                        try:
                            pdfmetrics.registerFont(TTFont(font_name, font_path))
                            print(f"✅ 成功注册Windows中文字体: {font_name}")
                            return font_name
                        except Exception as e:
                            continue
            
            # 最终备用方案 - 使用支持中文的内置字体
            print("⚠️ 未找到合适的中文字体，使用内置字体")
            return 'Helvetica'  # 至少保持基本显示
            
        except Exception as e:
            print(f"字体注册失败: {e}")
            return 'Helvetica'
    
    def create_box_label(self, canvas_obj, data, x, y, label_type='box', appearance='v1'):
        """
        创建单个盒标 - 支持两种外观样式
        
        Args:
            canvas_obj: ReportLab Canvas对象
            data: 标签数据字典
            x, y: 标签左下角坐标
            label_type: 标签类型 ('box'=盒标)
            appearance: 外观样式 ('v1'=原有样式, 'v2'=三行文本样式)
        """
        c = canvas_obj
        
        if appearance == 'v2':
            self.create_box_label_v2(c, data, x, y)
            return
        
        # 不绘制边框 - 标签无边框，纯文字显示
        # c.rect(x, y, self.LABEL_WIDTH, self.LABEL_HEIGHT)  # 注释掉边框
        
        # 内边距 - 调整为与PDF标签一致的宽松布局
        padding = 5 * mm  # 增加内边距使布局更宽松
        
        # 标签中心点
        center_x = x + self.LABEL_WIDTH / 2
        center_y = y + self.LABEL_HEIGHT / 2
        
        # 上方：主题文字 - 只提取英文部分
        raw_title = data.get('subject', data.get('B4', 'DEX\'S SIDEKICK'))
        
        # 智能提取英文部分
        if raw_title:
            import re
            # 先去掉开头的"-"符号（如果有）
            clean_title = raw_title.lstrip('-').strip()
            
            # 查找英文部分 - 匹配连续的英文字母、空格、撇号、感叹号等
            english_patterns = [
                r'[A-Z][A-Z\s\'!]*[A-Z!]',           # 大写字母开头结尾的英文短语
                r'[A-Z]+\'[A-Z\s]+[A-Z!]',           # 带撇号的英文 (如 DEX'S SIDEKICK)
                r'[A-Z]+[A-Z\s!]*',                  # 任何大写字母组合
                r'[A-Za-z][A-Za-z\s\'!]*[A-Za-z!]'  # 任何英文字母组合
            ]
            
            for pattern in english_patterns:
                match = re.search(pattern, clean_title)
                if match:
                    main_title = match.group().strip()
                    break
            else:
                # 如果没有匹配到，使用清理后的原标题
                main_title = clean_title if clean_title else 'TAG! YOU\'RE IT!'
        else:
            main_title = 'TAG! YOU\'RE IT!'
        
        # 调试输出
        print(f"原始标题: '{raw_title}' -> 处理后标题: '{main_title}'")
            
        # 重置绘制设置，确保文字正常渲染
        c.setFillColor(self.colors['black'])
        # 不设置描边，只使用填充模式绘制文字
        
        # 主题文字 - 强制使用简单内置字体避免渲染问题
        title_font_size = 18
        c.setFont('Helvetica-Bold', title_font_size)  # 直接使用内置字体，不用注册的复杂字体
        
        title_width = c.stringWidth(main_title, 'Helvetica-Bold', title_font_size)
        title_x = center_x - title_width / 2
        title_y = center_y + 18  # 向上移动更多，增加与编号的间距
        c.drawString(title_x, title_y, main_title)
        
        # 编号文字 - 同样使用简单内置字体
        product_code = data.get('start_number', data.get('B11', 'DSK01001'))
        code_font_size = 20  # 稍大于主题，匹配目标样式比例
        c.setFont('Helvetica-Bold', code_font_size)  # 直接使用内置字体
        
        code_width = c.stringWidth(product_code, 'Helvetica-Bold', code_font_size)
        code_x = center_x - code_width / 2
        code_y = center_y - 18  # 向下移动更多，增加与主题的间距
        c.drawString(code_x, code_y, product_code)
    
    def create_box_label_v2(self, canvas_obj, data, x, y):
        """
        创建外观2样式的盒标 - 三行文本布局
        Game title: XXX
        Ticket count: XXX  
        Serial: XXX
        """
        c = canvas_obj
        
        # 标签区域中心点
        center_x = x + self.LABEL_WIDTH / 2
        center_y = y + self.LABEL_HEIGHT / 2
        
        # 提取数据并进行数据映射
        # B4 -> Game title (提取英文部分)
        raw_theme = data.get('subject', data.get('B4', 'Lab forest'))
        game_title = self._extract_english_theme_v2(raw_theme)
        
        # Ticket count = 每盒张数 (不是总张数F4)
        ticket_count = data.get('min_box_count', data.get('box_count', 10))
        
        # B11 -> Serial (当前编号)
        serial = data.get('start_number', data.get('B11', 'LAF01001'))
        
        # 重置绘制设置
        c.setFillColor(self.colors['black'])
        
        # 字体设置 - 更大更粗的字体
        font_size = 18  # 增大字体大小
        
        # 根据标准图片精确定位三行文本
        # 精细调整：第一行稍向下，第二三行更紧密且更靠下
        title_y = y + self.LABEL_HEIGHT - 15 * mm     # Game title位置 - 向下移动3mm
        count_y = y + self.LABEL_HEIGHT - 36 * mm     # Ticket count位置 - 向下移动4mm
        serial_y = y + self.LABEL_HEIGHT - 46 * mm    # Serial位置 - 向下移动2mm，与第二行更紧密
        
        # 边距设置 - 确保左右都有适当边距
        left_margin = x + 4 * mm   # 左边距
        right_margin = 4 * mm      # 右边距（用于检查文本宽度）
        
        # 使用紧凑字符间距确保长文本能完整显示
        def draw_text_with_tight_spacing(canvas, x, y, text, font_size):
            """绘制紧凑间距的文本"""
            canvas.setFont('Helvetica-Bold', font_size)
            
            # 通过逐个字符绘制来控制间距
            current_x = x
            for char in text:
                canvas.drawString(current_x, y, char)
                char_width = canvas.stringWidth(char, 'Helvetica-Bold', font_size)
                current_x += char_width * 0.94  # 使用94%的间距，确保文本能完整显示
        
        # 第一行: Game title - 使用紧凑间距确保长文本显示完整
        title_text = f"Game title: {game_title}"
        draw_text_with_tight_spacing(c, left_margin, title_y, title_text, font_size)
        
        # 第二行: Ticket count - 使用紧凑间距
        count_text = f"Ticket count: {ticket_count}"
        draw_text_with_tight_spacing(c, left_margin, count_y, count_text, font_size)
        
        # 第三行: Serial - 使用紧凑间距
        serial_text = f"Serial: {serial}"
        draw_text_with_tight_spacing(c, left_margin, serial_y, serial_text, font_size)
        
        print(f"绘制外观2标签: Game='{game_title}', Ticket='{ticket_count}', Serial='{serial}'")
    
    def _extract_english_theme_v2(self, theme_text):
        """为外观2提取英文主题"""
        if not theme_text:
            return 'Lab forest'
        
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
        
        return clean_theme if clean_theme else 'Lab forest'
    
    def generate_labels_pdf(self, data_dict, quantities, output_path, label_prefix="", appearance='v1'):
        """
        生成多级标签PDF
        
        Args:
            data_dict: Excel数据字典，包含A4, B4, B11, F4的值
            quantities: 数量配置字典 {
                'min_box_count': 最小分盒张数,
                'box_per_inner_case': 盒/小箱,
                'inner_case_per_outer_case': 小箱/大箱
            }
            output_path: 输出文件路径
            label_prefix: 标签前缀用于文件命名
        """
        
        # 从Excel数据中提取信息
        # F4位置的值是总张数，用于计算盒标数量
        total_sheets = int(data_dict.get('F4', data_dict.get('total_quantity', 100)))
        min_box_count = quantities.get('min_box_count', 10)
        box_per_inner = quantities.get('box_per_inner_case', 5)
        inner_per_outer = quantities.get('inner_case_per_outer_case', 4)
        
        # 基于总张数计算需要的各级标签数量
        # 盒标数量 = 总张数 / 每盒最小张数 (向上取整)
        box_count = math.ceil(total_sheets / min_box_count)
        inner_case_count = math.ceil(box_count / box_per_inner)
        outer_case_count = math.ceil(inner_case_count / inner_per_outer)
        
        print(f"标签数量计算:")
        print(f"  总张数: {total_sheets}")
        print(f"  每盒张数: {min_box_count}")
        print(f"  盒标数量: {box_count}")
        print(f"  编号应该从 {data_dict.get('B11')} 开始，连续递增到第{box_count}个")
        
        # 生成三种标签文件
        output_dir = Path(output_path)
        # 从Excel数据中获取客户名称和主题
        customer_name = data_dict.get('A4', '默认客户')  # A4位置的客户名称
        theme = data_dict.get('B4', '默认主题')  # B4位置的主题
        
        # 创建文件夹
        folder_name = f"{customer_name}+{theme}+标签"
        label_folder = output_dir / folder_name
        label_folder.mkdir(exist_ok=True)
        
        # 准备标签数据 - 常规模版1简单数据
        label_data = {
            'customer_name': customer_name,  # A4
            'subject': theme,                # B4 - 主题
            'start_number': data_dict.get('B11', 'DSK01001'),  # B11 - 起始编号
            'total_quantity': total_sheets,  # F4 - 总张数
            'F4': data_dict.get('F4', total_sheets),  # 保留原始F4数据
            'B4': theme,  # 保留B4数据
            'min_box_count': min_box_count,  # 每盒张数 - 用于外观2的Ticket count
            'box_count': min_box_count  # 备用字段
        }
        
        # 生成盒标 - 根据外观选择生成不同样式
        appearance_suffix = "外观2" if appearance == 'v2' else ""
        box_label_path = label_folder / f"{customer_name}+{theme}+盒标{appearance_suffix}.pdf"
        self._generate_single_type_labels(
            label_data, box_count, str(box_label_path), 'box', appearance
        )
        
        print(f"✅ 生成盒标文件: {box_label_path.name}")
        
        return {
            'box_labels': str(box_label_path),
            'folder': str(label_folder),
            'count': box_count
        }
    
    def _generate_single_type_labels(self, data, count, output_path, label_type, appearance='v1'):
        """生成单一类型的标签PDF文件 - 90x50mm页面尺寸"""
        # 使用90x50mm作为页面尺寸
        page_size = (self.LABEL_WIDTH, self.LABEL_HEIGHT)  # 90x50mm
        c = canvas.Canvas(output_path, pagesize=page_size)
        
        # 设置PDF/X-3元数据（适用于CMYK打印）
        c.setTitle(f"盒标 - {data.get('subject', 'SIDEKICK')}")
        c.setAuthor("数据转PDF打印工具")
        c.setSubject("90x50mm盒标批量打印")
        c.setCreator("盒标生成工具 v2.0")
        c.setKeywords("盒标,标签,PDF/X,CMYK,打印")
        
        # PDF/X-3兼容性设置
        try:
            # 设置CMYK颜色空间
            c._doc.catalog.colorSpace = "/DeviceCMYK"
            # 添加PDF/X标识
            c._doc.catalog.GTS_PDFXVersion = "PDF/X-3:2002"
            c._doc.catalog.GTS_PDFXConformance = "PDF/X-3:2002"
        except:
            pass  # 如果ReportLab版本不支持则跳过
        
        # 90x50mm页面尺寸设置
        page_width, page_height = page_size  # 使用90x50mm页面
        
        # 由于页面就是标签尺寸，文字直接在页面中心
        labels_per_page = 1  # 每页1个标签
        
        # 文字在90x50mm页面中心
        start_x = 0  # 标签从页面左下角开始
        start_y = 0  # 标签从页面左下角开始
        
        print(f"页面布局: 90x50mm页面，每页{labels_per_page}个标签")
        
        current_label = 0
        
        for i in range(count):
            # 每个标签都居中显示
            x = start_x  # 水平居中
            y = start_y  # 垂直居中
            
            # 为每个标签准备数据
            label_data = data.copy()
            if 'start_number' in data:
                # 常规模版1：从开始号开始每次加1
                base_number = str(data['start_number'])
                label_data['start_number'] = self._generate_simple_increment(base_number, i)
                print(f"盒标 {i+1}: 编号 {label_data['start_number']} (第{i+1}页)")
            
            # 创建标签
            self.create_box_label(c, label_data, x, y, label_type, appearance)
            
            # 每个标签后都换页（除了最后一个）
            if i < count - 1:
                c.showPage()
        
        # 保存PDF
        c.save()
    
    def _generate_simple_increment(self, base_number, box_index):
        """
        常规模版1：从开始号开始每次加1
        
        Args:
            base_number: 基础编号 (如: DSK01001)
            box_index: 盒子索引（从0开始）
        
        Returns:
            str: 递增后的编号
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
                # 简单递增：每个盒标编号 = 开始号 + 盒标索引
                current_number = start_num + box_index
                # 保持原数字部分的位数
                width = len(number_part)
                result = f"{prefix_part}{current_number:0{width}d}"
                return result
            else:
                # 如果无法解析数字，使用简单递增
                return f"{base_number}_{box_index+1:03d}"
                
        except Exception as e:
            print(f"简单递增编号生成失败: {e}")
            return f"{base_number}_{box_index+1:03d}"


class BoxLabelDataExtractor:
    """盒标数据提取器 - 专门处理Excel特定位置的数据"""
    
    @staticmethod
    def extract_from_excel(excel_reader, sheet_name=None):
        """
        从Excel提取盒标相关数据
        
        Args:
            excel_reader: ExcelReader实例
            sheet_name: 工作表名称，如果为None则使用第一个工作表
        
        Returns:
            dict: 包含A4, B4, B11, F4位置数据的字典
        """
        try:
            # 读取原始Excel数据（不使用pandas的默认处理）
            import openpyxl
            
            workbook = openpyxl.load_workbook(excel_reader.file_path)
            if sheet_name:
                worksheet = workbook[sheet_name]
            else:
                worksheet = workbook.active
            
            # 提取指定单元格数据
            extracted_data = {
                'customer_name': worksheet['A4'].value,  # 客户名称编码
                'subject': worksheet['B4'].value,        # 主题
                'start_number': worksheet['B11'].value,  # 开始号
                'total_quantity': worksheet['F4'].value  # 总张数
            }
            
            # 清理数据 - 移除None值并转换为字符串
            for key, value in extracted_data.items():
                if value is None:
                    extracted_data[key] = ""
                else:
                    extracted_data[key] = str(value).strip()
            
            # 数据验证
            if not extracted_data['total_quantity'].isdigit():
                try:
                    extracted_data['total_quantity'] = int(float(extracted_data['total_quantity']))
                except:
                    extracted_data['total_quantity'] = 100  # 默认值
            
            workbook.close()
            return extracted_data
            
        except Exception as e:
            raise Exception(f"数据提取失败: {str(e)}")