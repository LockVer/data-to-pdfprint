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
        
        # 上方：主题文字 - 使用新的搜索方式获取主题
        # 如果data中包含excel_data，则使用搜索方式，否则使用原有方式
        if 'excel_data' in data:
            main_title = self._search_label_name_data(data['excel_data'])
        else:
            # 原有的主题提取方式（保持向后兼容）
            raw_title = data.get('subject', data.get('B4', 'DEX\'S SIDEKICK'))
            main_title = self._extract_english_theme(raw_title)
        
        # 调试输出
        print(f"盒标主题: '{main_title}'")
            
        # 重置绘制设置，确保文字正常渲染
        c.setFillColor(self.colors['black'])
        # 不设置描边，只使用填充模式绘制文字
        
        # 主题文字 - 使用粗体绘制方法，与分合模版保持一致
        title_font_size = 18  # 与分合模版保持一致
        
        # 主题区域设置
        theme_max_width = self.LABEL_WIDTH - 6 * mm  # 左右各留3mm边距
        theme_max_height = self.LABEL_HEIGHT * 0.6   # 给主题文字更多空间，60%的高度
        theme_x = x + 3 * mm  # 左边距
        theme_y = y + self.LABEL_HEIGHT - 9 * mm  # 从标签顶部向下9mm开始
        
        # 使用多行粗体绘制，支持自动换行和居中对齐
        self._draw_bold_multiline_text(
            c, main_title, theme_x, theme_y,
            theme_max_width, theme_max_height, 
            self.chinese_font, title_font_size,
            align='center'  # 居中对齐
        )
        
        # 编号文字 - 使用粗体绘制方法，与分合模版保持一致
        product_code = data.get('start_number', data.get('B11', 'DSK01001'))
        code_font_size = 18  # 与主题保持一致，统一为18pt
        
        # 编号区域设置
        number_area_width = self.LABEL_WIDTH * 0.9  # 编号区域宽度
        number_start_x = x + (self.LABEL_WIDTH - number_area_width) / 2
        number_start_y = y + self.LABEL_HEIGHT * 0.25  # 更靠下的位置，增加与主题的间距
        
        self._draw_bold_single_line(
            c, product_code, number_start_x, number_start_y,
            number_area_width, self.chinese_font, code_font_size,
            align='center'
        )
    
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
        # 使用新的搜索方式获取主题
        if 'excel_data' in data:
            game_title = self._search_label_name_data(data['excel_data'])
        else:
            # 原有的主题提取方式（保持向后兼容）
            raw_theme = data.get('subject', data.get('B4', 'Lab forest'))
            game_title = self._extract_english_theme_v2(raw_theme)
        
        # Ticket count = 每盒张数 (不是总张数F4)
        ticket_count = data.get('min_box_count', data.get('box_count', 10))
        
        # B11 -> Serial (当前编号)
        serial = data.get('start_number', data.get('B11', 'LAF01001'))
        
        # 重置绘制设置
        c.setFillColor(self.colors['black'])
        
        # 字体设置 - 与分合模版保持一致
        font_size = 18  # 与分合模版保持一致
        
        # 根据标准图片精确定位三行文本
        # 精细调整：第一行稍向下，第二三行更紧密且更靠下
        title_y = y + self.LABEL_HEIGHT - 15 * mm     # Game title位置 - 向下移动3mm
        count_y = y + self.LABEL_HEIGHT - 36 * mm     # Ticket count位置 - 向下移动4mm
        serial_y = y + self.LABEL_HEIGHT - 46 * mm    # Serial位置 - 向下移动2mm，与第二行更紧密
        
        # 边距设置 - 确保左右都有适当边距
        left_margin = x + 4 * mm   # 左边距
        right_margin = 4 * mm      # 右边距（用于检查文本宽度）
        
        # 第一行: Game title - 使用粗体绘制方法
        title_text = f"Game title: {game_title}"
        self._draw_bold_text(c, title_text, left_margin, title_y, self.chinese_font, font_size)
        
        # 第二行: Ticket count - 使用粗体绘制方法
        count_text = f"Ticket count: {ticket_count}"
        self._draw_bold_text(c, count_text, left_margin, count_y, self.chinese_font, font_size)
        
        # 第三行: Serial - 使用粗体绘制方法
        serial_text = f"Serial: {serial}"
        self._draw_bold_text(c, serial_text, left_margin, serial_y, self.chinese_font, font_size)
        
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
    
    def _extract_english_theme(self, theme_text):
        """提取英文主题 - 原有方式，保持向后兼容"""
        if not theme_text:
            return 'DEX\'S SIDEKICK'
        
        import re
        # 先去掉开头的"-"符号（如果有）
        clean_title = theme_text.lstrip('-').strip()
        
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
                return match.group().strip()
        
        # 如果没有匹配到，使用清理后的原标题
        return clean_title if clean_title else 'TAG! YOU\'RE IT!'
    
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
    
    def _draw_bold_multiline_text(self, canvas_obj, text, x, y, max_width, max_height, font_name, font_size, align='left'):
        """
        绘制支持自动换行的粗体多行文本（通过重复绘制实现粗体效果）
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
        
        line_height = font_size * 1.2
        
        # 禁用自动字体调整，保持设定的大字体效果
        start_y = y - font_size
        
        # 粗体效果的偏移量 - 减小偏移量避免重影
        bold_offsets = [
            (0, 0),      # 原始位置
            (0.3, 0),    # 右偏移
            (0, 0.3),    # 上偏移  
            (0.3, 0.3),  # 右上偏移
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
    
    def _draw_bold_single_line(self, canvas_obj, text, x, y, max_width, font_name, font_size, align='left'):
        """
        绘制单行粗体文本（通过重复绘制实现粗体效果）
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
        
        # 粗体效果的偏移量 - 减小偏移量避免重影
        bold_offsets = [
            (0, 0),      # 原始位置
            (0.3, 0),    # 右偏移
            (0, 0.3),    # 上偏移  
            (0.3, 0.3),  # 右上偏移
        ]
        
        # 多次绘制实现粗体效果
        for offset_x, offset_y in bold_offsets:
            c.drawString(base_x + offset_x, y + offset_y, text)
    
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
            'box_count': min_box_count,  # 备用字段
            'excel_data': data_dict  # 添加完整的Excel数据，用于主题搜索
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