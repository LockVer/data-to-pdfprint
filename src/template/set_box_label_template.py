"""
套盒盒标模板系统

专门用于生成套盒盒标PDF，支持特殊的箱号-盒号编号格式
编号规则：基础编号-盒在箱内序号 (如：MOP01001-01, MOP01002-01)
每满指定盒数进入下一箱，箱内盒号重新从01开始，箱号加1
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

# 导入统一的字体工具
try:
    from .font_utils import get_chinese_font
except ImportError:
    def get_chinese_font():
        return 'Helvetica' 

class SetBoxLabelTemplate:
    """套盒盒标模板类"""
    
    # 标签尺寸 (90x50mm)
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
    
    def create_set_box_label(self, canvas_obj, data, x, y):
        """
        创建单个套盒盒标
        
        Args:
            canvas_obj: ReportLab Canvas对象
            data: 标签数据字典
            x, y: 标签左下角坐标
        """
        c = canvas_obj
        
        # 标签中心点
        center_x = x + self.LABEL_WIDTH / 2
        center_y = y + self.LABEL_HEIGHT / 2
        
        # 上方：主题文字 - 使用和分盒模板相同的搜索逻辑
        main_title = self._search_label_name_data(data)
        if not main_title:
            # 备选方案：使用B4数据
            main_title = data.get('subject', data.get('B4', 'TAB STREET DRAMA'))
        
        # 调试输出
        print(f"套盒模板主题获取结果: '{main_title}'")
            
        # 重置绘制设置，确保文字正常渲染
        c.setFillColor(self.colors['black'])
        
        # 主题文字 - 强制使用简单内置字体避免渲染问题
        title_font_size = 18
        c.setFont('Helvetica-Bold', title_font_size)
        
        title_width = c.stringWidth(main_title, 'Helvetica-Bold', title_font_size)
        title_x = center_x - title_width / 2
        title_y = center_y + 18  # 向上移动更多，增加与编号的间距
        # 使用粗体绘制方法绘制主题
        self._draw_bold_text(c, main_title, title_x, title_y, 'Helvetica-Bold', title_font_size)
        
        # 套盒编号文字 - 使用箱号-盒号格式
        set_box_code = data.get('set_box_number', 'MOP01001-01')
        code_font_size = 20  # 稍大于主题，匹配目标样式比例
        c.setFont('Helvetica-Bold', code_font_size)
        
        code_width = c.stringWidth(set_box_code, 'Helvetica-Bold', code_font_size)
        code_x = center_x - code_width / 2
        code_y = center_y - 18  # 向下移动更多，增加与主题的间距
        # 使用粗体绘制方法绘制编号
        self._draw_bold_text(c, set_box_code, code_x, code_y, 'Helvetica-Bold', code_font_size)
    
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
        
        # 如果没找到"标签名称"关键字，返回None
        print(f"⚠️  未找到标签名称关键字")
        return None
    
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

    def _generate_set_box_number(self, base_number, box_index, boxes_per_set):
        """
        生成套盒编号 - 套号-盒号格式
        
        Args:
            base_number: 基础编号 (如: MOP01001)
            box_index: 盒子索引（从0开始）
            boxes_per_set: 几盒为一套
        
        Returns:
            str: 套盒编号 (如: MOP01001-01, MOP01001-02, MOP01001-03, MOP01002-01...)
        """
        try:
            # 首先清理基础编号，去掉可能存在的"-XX"后缀
            clean_base_number = base_number
            if '-' in base_number:
                # 如果基础编号包含"-"，取第一部分作为真正的基础编号
                clean_base_number = base_number.split('-')[0]
                print(f"  - 清理基础编号: '{base_number}' -> '{clean_base_number}'")
            
            # 计算套号和盒在套内的序号
            set_index = box_index // boxes_per_set  # 套索引（从0开始）
            box_in_set = box_index % boxes_per_set + 1  # 盒在套内序号（从1开始）
            
            # 提取前缀和数字部分
            prefix_part = ''
            number_part = ''
            
            # 从后往前找连续的数字
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
                
                print(f"🎯🎯🎯 套盒编号生成详细:")
                print(f"  - 盒索引: {box_index} (第{box_index+1}盒)")
                print(f"  - 几盒为一套: {boxes_per_set}")
                print(f"  - 套索引计算: {box_index} ÷ {boxes_per_set} = {set_index}")
                print(f"  - 套内盒序号计算: {box_index} % {boxes_per_set} + 1 = {box_in_set}")
                print(f"  - 清理后基础编号: {clean_base_number} -> 套号: {set_part}")
                print(f"  - 🔥 最终编号: {result} 🔥")
                return result
            else:
                # 如果无法解析数字，使用简单格式
                set_part = f"{base_number}_{set_index+1:03d}"
                return f"{set_part}-{box_in_set:02d}"
                
        except Exception as e:
            print(f"套盒编号生成失败: {e}")
            return f"{base_number}_SET{set_index+1:03d}-{box_in_set:02d}"

    def generate_set_box_labels_pdf(self, data_dict, quantities, output_path, label_prefix=""):
        """
        生成套盒盒标PDF
        
        Args:
            data_dict: Excel数据字典，包含A4, B4, B11, F4的值
            quantities: 数量配置字典 {
                'min_box_count': 每盒张数,
                'boxes_per_set': 几盒为一套,
                'boxes_per_inner_case': 几盒入一小箱,
                'sets_per_outer_case': 几套入一大箱
            }
            output_path: 输出文件路径
            label_prefix: 标签前缀用于文件命名
        """
        
        # 从Excel数据中提取信息
        # F4位置的值是总张数，用于计算套数
        total_sheets = int(data_dict.get('F4', data_dict.get('total_quantity', 100)))
        # 获取每套张数，优先级：min_set_count > min_box_count*3 > 默认30
        if 'min_set_count' in quantities:
            min_set_count = quantities['min_set_count']
        elif 'min_box_count' in quantities:
            # 从旧参数转换：每盒张数 * 默认每套盒数(3) = 每套张数
            min_set_count = quantities['min_box_count'] * quantities.get('boxes_per_set', 3)
        else:
            min_set_count = 30  # 默认每套30张
        # 新的套盒参数结构
        boxes_per_set = quantities.get('boxes_per_set', 3)  # 几盒为一套
        boxes_per_inner_case = quantities.get('boxes_per_inner_case', 6)  # 几盒入一小箱
        sets_per_outer_case = quantities.get('sets_per_outer_case', 2)  # 几套入一大箱
        
        # 基于总张数计算需要的套数和盒标数量
        # 套数 = 总张数 / 每套张数 (向上取整)
        set_count = math.ceil(total_sheets / min_set_count)
        # 盒标数量 = 套数 × 每套盒数
        box_count = set_count * boxes_per_set
        
        print(f"=" * 80)
        print(f"🎯🎯🎯 正在使用套盒模版生成标签！🎯🎯🎯")
        print(f"=" * 80)
        print(f"套盒标签数量计算:")
        print(f"  总张数: {total_sheets}")
        print(f"  每套张数: {min_set_count}")
        print(f"  几盒为一套: {boxes_per_set}")
        print(f"  几盒入一小箱: {boxes_per_inner_case}")
        print(f"  几套入一大箱: {sets_per_outer_case}")
        print(f"  套数: {set_count}")
        print(f"  盒标数量: {box_count}")
        print(f"  编号应该从 {data_dict.get('B11')} 开始，按照套盒逻辑递增")
        print(f"  编号格式: JAW01001-01 (第1套第1盒), JAW01001-02 (第1套第2盒)...")
        print(f"=" * 80)
        
        # 生成套盒标签文件
        output_dir = Path(output_path)
        # 从Excel数据中获取客户名称和主题
        customer_name = data_dict.get('A4', '默认客户')  # A4位置的客户名称
        theme = data_dict.get('B4', '默认主题')  # B4位置的主题
        
        # 创建文件夹 - 清理文件名中的非法字符
        import re
        # 清理客户名称和主题中的非法字符
        clean_customer_name = re.sub(r'[<>:"/\\|?*]', '_', str(customer_name))
        clean_theme = re.sub(r'[<>:"/\\|?*]', '_', str(theme))
        
        folder_name = f"{clean_customer_name}+{clean_theme}+标签"
        label_folder = output_dir / folder_name
        
        try:
            label_folder.mkdir(exist_ok=True)
            print(f"✅ 成功创建输出文件夹: {label_folder}")
        except PermissionError as e:
            raise Exception(f"权限错误：无法在选择的目录中创建文件夹。请选择一个有写权限的目录。\n错误详情: {e}")
        except OSError as e:
            if "Read-only file system" in str(e):
                raise Exception(f"文件系统错误：选择的目录是只读的，无法创建文件。请选择一个可写的目录。\n路径: {output_dir}")
            else:
                raise Exception(f"创建文件夹失败: {e}\n路径: {label_folder}")
        
        # 准备标签数据 - 套盒模版数据，包含完整的Excel数据用于搜索
        label_data = data_dict.copy()  # 保留完整的Excel数据
        label_data.update({
            'customer_name': customer_name,  # A4
            'subject': theme,                # B4 - 主题
            'start_number': data_dict.get('B11', 'MOP01001'),  # B11 - 起始编号
            'total_quantity': total_sheets,  # F4 - 总张数
            'F4': data_dict.get('F4', total_sheets),  # 保留原始F4数据
            'B4': theme,  # 保留B4数据
            'min_set_count': min_set_count,  # 每套张数
            'set_count': set_count,  # 套数
            'boxes_per_set': boxes_per_set,  # 几盒为一套
            'boxes_per_inner_case': boxes_per_inner_case,  # 几盒入一小箱
            'sets_per_outer_case': sets_per_outer_case  # 几套入一大箱
        })
        
        # 生成套盒盒标
        set_box_label_path = label_folder / f"{customer_name}+{theme}+套盒盒标.pdf"
        self._generate_set_box_labels(
            label_data, box_count, str(set_box_label_path), boxes_per_set
        )
        
        print(f"✅ 生成套盒盒标文件: {set_box_label_path.name}")
        
        # 生成套盒小箱标和大箱标
        result = {
            'set_box_labels': str(set_box_label_path),
            'folder': str(label_folder),
            'count': box_count
        }
        
        try:
            # 生成套盒小箱标
            from .set_box_inner_case_template import SetBoxInnerCaseTemplate
            inner_case_template = SetBoxInnerCaseTemplate()
            inner_case_result = inner_case_template.generate_set_box_inner_case_labels_pdf(
                data_dict, quantities, output_dir
            )
            print(f"✅ 生成套盒小箱标文件: {Path(inner_case_result['set_box_inner_case_labels']).name}")
            
            result['set_box_inner_case_labels'] = inner_case_result['set_box_inner_case_labels']
            result['inner_case_count'] = inner_case_result['count']
            
        except Exception as e:
            print(f"⚠️ 套盒小箱标生成失败: {e}")
        
        try:
            # 生成套盒大箱标
            from .set_box_outer_case_template import SetBoxOuterCaseTemplate
            outer_case_template = SetBoxOuterCaseTemplate()
            outer_case_result = outer_case_template.generate_set_box_outer_case_labels_pdf(
                data_dict, quantities, output_dir
            )
            print(f"✅ 生成套盒大箱标文件: {Path(outer_case_result['set_box_outer_case_labels']).name}")
            
            result['set_box_outer_case_labels'] = outer_case_result['set_box_outer_case_labels']
            result['outer_case_count'] = outer_case_result['count']
            
        except Exception as e:
            print(f"⚠️ 套盒大箱标生成失败: {e}")
        
        return result
    
    def _generate_set_box_labels(self, data, count, output_path, boxes_per_set):
        """生成套盒盒标PDF文件 - 90x50mm页面尺寸"""
        # 使用90x50mm作为页面尺寸
        page_size = (self.LABEL_WIDTH, self.LABEL_HEIGHT)  # 90x50mm
        c = canvas.Canvas(output_path, pagesize=page_size)
        
        # 设置PDF/X-3元数据（适用于CMYK打印）
        c.setTitle(f"套盒盒标 - {data.get('subject', 'SET BOX')}")
        c.setAuthor("数据转PDF打印工具")
        c.setSubject("90x50mm套盒盒标批量打印")
        c.setCreator("套盒盒标生成工具 v1.0")
        c.setKeywords("套盒盒标,标签,PDF/X,CMYK,打印")
        
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
        
        for i in range(count):
            # 每个标签都居中显示
            x = start_x  # 水平居中
            y = start_y  # 垂直居中
            
            # 为每个标签准备数据
            label_data = data.copy()
            if 'start_number' in data:
                # 套盒模版：生成箱号-盒号格式的编号
                base_number = str(data['start_number'])
                set_box_number = self._generate_set_box_number(base_number, i, boxes_per_set)
                label_data['set_box_number'] = set_box_number
                print(f"套盒盒标 {i+1}: 编号 {set_box_number} (第{i+1}页)")
            
            # 创建标签
            self.create_set_box_label(c, label_data, x, y)
            
            # 每个标签后都换页（除了最后一个）
            if i < count - 1:
                c.showPage()
        
        # 保存PDF
        c.save()


class SetBoxLabelDataExtractor:
    """套盒盒标数据提取器 - 专门处理Excel特定位置的数据"""
    
    @staticmethod
    def extract_from_excel(excel_reader, sheet_name=None):
        """
        从Excel提取套盒盒标相关数据
        
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
            raise Exception(f"套盒数据提取失败: {str(e)}")