import os
import re
import math
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from data.excel_reader import ExcelReader

# 指定字体路径
font_path_msyh = os.path.join(os.path.dirname(__file__), 'fonts', 'msyh.ttf')
font_path_simsun = os.path.join(os.path.dirname(__file__), 'fonts', 'SimSun.ttf')

# 注册字体
pdfmetrics.registerFont(TTFont('MicrosoftYaHei', font_path_msyh))
pdfmetrics.registerFont(TTFont('SimSun', font_path_simsun))

def extract_english_filename(file_path, theme=None):
    """
    从文件路径中提取英文部分作为文件名，或使用theme字段
    
    Args:
        file_path: 原始文件路径
        theme: Excel中提取的theme字段（优先使用）
    
    Returns:
        清理后的英文文件名
    """
    if theme and theme.strip():
        # 优先使用theme字段，清理特殊字符
        clean_theme = re.sub(r'[^\w\s]', '', theme)  # 移除特殊字符，保留字母数字和空格
        clean_theme = re.sub(r'\s+', '_', clean_theme.strip())  # 空格替换为下划线
        return clean_theme
    
    # 从文件路径提取英文部分
    filename = os.path.basename(file_path)
    # 移除文件扩展名
    name_without_ext = os.path.splitext(filename)[0]
    
    # 删除中文字符（Unicode范围）
    english_only = re.sub(r'[\u4e00-\u9fff]', '', name_without_ext)
    # 删除连字符
    english_only = english_only.replace('-', '')
    # 删除特殊字符，只保留字母数字和空格
    english_only = re.sub(r'[^\w\s]', '', english_only)
    # 清理多余空格并替换为下划线
    english_only = re.sub(r'\s+', '_', english_only.strip())
    
    # 如果没有英文内容，使用默认名称
    if not english_only:
        english_only = "case_output"
    
    return english_only

class CaseTemplate:
    def __init__(self, width_pt, height_pt, style_refs, pages):
        self.width_pt = width_pt
        self.height_pt = height_pt
        self.pages = pages
        self.style_refs = style_refs

class CasePage:
    def __init__(self, index, elements, index_range=None):
        self.index = index
        self.elements = elements
        self.index_range = index_range

class CaseElement:
    def __init__(self, role, content, content_template=None, seq=None):
        self.role = role
        self.content = content
        self.content_template = content_template
        self.seq = seq

class CaseSequence:
    def __init__(self, start, end, pad=1):
        self.start = start
        self.end = end
        self.pad = pad

class CaseStyle:
    def __init__(self, font_family, font_weight="normal", font_size=12, x_align="left", y_align="top"):
        self.font_family = font_family
        self.font_weight = font_weight
        self.font_size = font_size
        self.x_align = x_align
        self.y_align = y_align

def create_case_template_data(excel_variables, additional_inputs=None, template_mode="regular_single"):
    """
    根据Excel数据和额外输入值创建箱标template_data
    
    Args:
        excel_variables: 从Excel提取的变量 {
            'customer_code': str,  # A3: 客户名称编码 → Remark
            'theme': str,          # B3: 主题 → Theme  
            'start_number': str,   # B10: 开始号 → 用于序列号
            'total_sheets': int    # F3: 总张数 → 用于计算
        }
        additional_inputs: 额外输入值 (可选) {
            'sheets_per_box': int,           # 每盒张数
            'boxes_per_small_case': int,     # 每小箱盒数  
            'small_cases_per_large_case': int # 每大箱小箱数
        }
        template_mode: 箱标模板模式 {
            'regular_single': 单级常规模式
            'multi_separate': 多级分盒模式 
            'multi_set': 多级分套模式
        }
    
    Returns:
        CaseTemplate对象 (包含小箱标和大箱标)
    """
    
    # 提取基础数据
    customer_code = excel_variables.get('customer_code', '')
    theme = excel_variables.get('theme', '') 
    start_number = excel_variables.get('start_number', '1')
    total_cards = excel_variables.get('total_sheets', 1)
    
    # 提取额外参数
    if additional_inputs:
        sheets_per_box = additional_inputs.get('sheets_per_box', 2850)
        boxes_per_small_case = additional_inputs.get('boxes_per_small_case', 1)
        small_cases_per_large_case = additional_inputs.get('small_cases_per_large_case', 2)
    else:
        sheets_per_box = 2850
        boxes_per_small_case = 1  
        small_cases_per_large_case = 2
    
    # 根据模式调整计算逻辑
    if template_mode == "multi_set":
        # 套盒模式：一套=一小箱，sheets_per_box 实际是每套张数
        cards_per_set = sheets_per_box  # 每套张数（用户输入）
        sets_per_large_case = small_cases_per_large_case  # 每大箱套数（用户输入）
        
        # 从boxes_per_small_case参数推断每套盒数（从序列号格式01-06可知是6盒）
        boxes_per_set = boxes_per_small_case  # 每套盒数
        cards_per_box = cards_per_set // boxes_per_set if boxes_per_set > 0 else cards_per_set  # 每盒张数
        
        # 重新计算数量：套盒模式下，小箱=套
        total_sets = math.ceil(total_cards / cards_per_set) if cards_per_set > 0 else 0
        total_small_cases = total_sets  # 总小箱数=总套数（一套一小箱）
        total_large_cases = math.ceil(total_sets / sets_per_large_case) if sets_per_large_case > 0 else 0
        total_boxes = total_sets * boxes_per_set  # 总盒数
    else:
        # 其他模式：按原逻辑
        cards_per_box = sheets_per_box
        total_boxes = math.ceil(total_cards / cards_per_box)
        total_small_cases = math.ceil(total_boxes / boxes_per_small_case)
        total_large_cases = math.ceil(total_small_cases / small_cases_per_large_case)
    
    # 解析开始号前缀和数字
    match = re.search(r'^([A-Za-z]+)(\d+)', start_number)
    if match:
        serial_prefix = match.group(1)
        start_num = int(match.group(2))
    else:
        serial_prefix = customer_code
        start_num = 1
    
    pages = []
    
    if template_mode == "regular_single":
        # =============================================
        # 单级常规模式：小箱标(单盒) + 大箱标(多盒)
        # =============================================
        
        # 小箱标: 单个盒子 (2850PCS, LAN01001-LAN01001, 1/20)
        for case_idx in range(total_small_cases):
            box_num = start_num + case_idx
            box_serial = f"{serial_prefix}{box_num:05d}"
            
            pages.append(CasePage(index=case_idx*2 + 1, elements=[
                CaseElement(role="item", content="Paper Cards"),
                CaseElement(role="theme", content=theme),
                CaseElement(role="quantity", content=f"{cards_per_box}PCS"),
                CaseElement(role="serial_range", content=f"{box_serial}-{box_serial}"),
                CaseElement(role="carton_no", content=f"{case_idx + 1}/{total_small_cases}"),
                CaseElement(role="remark", content=customer_code)
            ]))
        
        # 大箱标: 多个盒子范围 (5700PCS, LAN01017-LAN01018, 9/10)
        for large_case_idx in range(total_large_cases):
            start_small_case = large_case_idx * small_cases_per_large_case
            end_small_case = min(start_small_case + small_cases_per_large_case - 1, total_small_cases - 1)
            
            start_box_num = start_num + start_small_case
            end_box_num = start_num + end_small_case
            
            large_quantity = (end_small_case - start_small_case + 1) * cards_per_box
            
            pages.append(CasePage(index=large_case_idx*2 + 2, elements=[
                CaseElement(role="item", content="Paper Cards"),
                CaseElement(role="theme", content=theme),
                CaseElement(role="quantity", content=f"{large_quantity}PCS"),
                CaseElement(role="serial_range", content=f"{serial_prefix}{start_box_num:05d}-{serial_prefix}{end_box_num:05d}"),
                CaseElement(role="carton_no", content=f"{large_case_idx + 1}/{total_large_cases}"),
                CaseElement(role="remark", content=customer_code)
            ]))
            
    elif template_mode == "multi_separate":
        # =============================================
        # 多级分盒模式：小箱标(单多级盒) + 大箱标(多多级盒)
        # 按照新的分盒模式规律生成
        # =============================================
        
        # 小箱标: 双层循环，多级格式
        for large_case_idx in range(total_large_cases):  # 大箱数量循环
            for small_case_idx in range(small_cases_per_large_case):  # 大箱内小箱数量循环
                if large_case_idx * small_cases_per_large_case + small_case_idx >= total_small_cases:
                    break
                
                # 父级编号（父级标号与子级编号相同）
                parent_num = large_case_idx + 1  # 大箱数量-循环递增
                child_num = small_case_idx + 1   # 小箱数量-循环递增
                
                # quantity: 盒张数 * 每小箱盒数（写死1）
                small_quantity = cards_per_box * 1  # 每小箱盒数写死为1
                
                # serial: 开始号和结束号相同，多级格式
                parent_serial = f"{serial_prefix}{parent_num + start_num - 1:05d}"
                child_serial = f"{child_num:02d}"
                multi_serial = f"{parent_serial}-{child_serial}"
                
                # carton_no: 双层循环格式
                carton_no = f"{parent_num}-{child_num}"
                
                pages.append(CasePage(index=large_case_idx*small_cases_per_large_case + small_case_idx + 1, elements=[
                    CaseElement(role="item", content="Paper Cards"),
                    CaseElement(role="theme", content=f"{theme} (chip)"),
                    CaseElement(role="quantity", content=f"{small_quantity}PCS"),
                    CaseElement(role="serial_range", content=f"{multi_serial}-{multi_serial}"),
                    CaseElement(role="carton_no", content=carton_no),
                    CaseElement(role="remark", content=customer_code)
                ]))
        
        # 大箱标: 跨范围，多级格式
        for large_case_idx in range(total_large_cases):
            # quantity: 大箱内小箱数量 * 盒张数 * 每小箱盒数（写死1）
            large_quantity = small_cases_per_large_case * cards_per_box * 1
            
            # serial: 开始号和结束号不同（跨范围）
            current_large_case_num = large_case_idx + 1
            parent_serial = f"{serial_prefix}{current_large_case_num + start_num - 1:05d}"
            
            # 开始号: 父级编号-子级编号（01固定开始）
            start_serial = f"{parent_serial}-01"
            
            # 结束号: 父级编号-子级编号（大箱内小箱数量最大值）
            end_serial = f"{parent_serial}-{small_cases_per_large_case:02d}"
            
            # carton_no: 当前大箱数
            carton_no = f"{current_large_case_num}"
            
            pages.append(CasePage(index=len(pages) + 1, elements=[
                CaseElement(role="item", content="Paper Cards"),
                CaseElement(role="theme", content=f"{theme} (chip)"),
                CaseElement(role="quantity", content=f"{large_quantity}PCS"),
                CaseElement(role="serial_range", content=f"{start_serial}-{end_serial}"),
                CaseElement(role="carton_no", content=carton_no),
                CaseElement(role="remark", content=customer_code)
            ]))
            
    elif template_mode == "multi_set":
        # =============================================
        # 多级分套模式：小箱标(套内多级) + 大箱标(跨套多级)
        # =============================================
        
        # 小箱标: 套内多级范围 (3780PCS, JAW01001-01-JAW01001-06, 01)
        for case_idx in range(total_small_cases):
            set_num = start_num + case_idx
            start_serial = f"{serial_prefix}{set_num:05d}-01"
            end_serial = f"{serial_prefix}{set_num:05d}-{boxes_per_set:02d}"
            
            # 套盒模式：一套=一小箱，所以就是每套张数
            small_quantity = cards_per_set
            
            pages.append(CasePage(index=case_idx*2 + 1, elements=[
                CaseElement(role="item", content="Paper Cards"),
                CaseElement(role="theme", content=theme),
                CaseElement(role="quantity", content=f"{small_quantity}PCS"),
                CaseElement(role="serial_range", content=f"{start_serial}-{end_serial}"),
                CaseElement(role="carton_no", content=f"{case_idx + 1:02d}"),
                CaseElement(role="remark", content=customer_code)
            ]))
        
        # 大箱标: 跨套多级范围 (7560PCS, JAW01037-01-JAW01038-06, 37-38)
        for large_case_idx in range(total_large_cases):
            start_case = large_case_idx * small_cases_per_large_case
            end_case = min(start_case + small_cases_per_large_case - 1, total_small_cases - 1)
            
            start_set_num = start_num + start_case
            end_set_num = start_num + end_case
            
            start_serial = f"{serial_prefix}{start_set_num:05d}-01"
            end_serial = f"{serial_prefix}{end_set_num:05d}-{boxes_per_set:02d}"
            
            # 套盒模式：该大箱包含的套数 * 每套张数
            large_quantity = (end_case - start_case + 1) * cards_per_set
            
            pages.append(CasePage(index=large_case_idx*2 + 2, elements=[
                CaseElement(role="item", content="Paper Cards"),
                CaseElement(role="theme", content=theme),
                CaseElement(role="quantity", content=f"{large_quantity}PCS"),
                CaseElement(role="serial_range", content=f"{start_serial}-{end_serial}"),
                CaseElement(role="carton_no", content=f"{start_set_num}-{end_set_num}"),
                CaseElement(role="remark", content=customer_code)
            ]))
    
    # 箱标使用90mm×50mm尺寸 (1mm = 2.834646 points)
    width_mm = 90
    height_mm = 50
    width_pt = width_mm * 2.834646
    height_pt = height_mm * 2.834646
    
    return CaseTemplate(
        width_pt=width_pt,
        height_pt=height_pt, 
        style_refs=["item", "theme", "quantity", "serial_range", "carton_no", "remark"],
        pages=pages
    )

def render_case_template_to_pdf(template, output_path):
    """
    将箱标模板渲染为PDF
    
    Args:
        template: CaseTemplate对象
        output_path: 输出PDF文件路径
    """
    c = canvas.Canvas(output_path, pagesize=(template.width_pt, template.height_pt))
    page_width, page_height = template.width_pt, template.height_pt
    
    # 箱标样式设置 - 表格形式布局
    case_style_set = {
        "item": CaseStyle(font_family="MicrosoftYaHei", font_weight="bold", font_size=20, x_align="center"),
        "theme": CaseStyle(font_family="MicrosoftYaHei", font_weight="bold", font_size=18, x_align="center"),
        "quantity": CaseStyle(font_family="MicrosoftYaHei", font_weight="bold", font_size=16, x_align="center"),
        "serial_range": CaseStyle(font_family="MicrosoftYaHei", font_size=14, x_align="center"),
        "carton_no": CaseStyle(font_family="MicrosoftYaHei", font_weight="bold", font_size=16, x_align="center"),
        "remark": CaseStyle(font_family="MicrosoftYaHei", font_size=14, x_align="center")
    }
    
    for page in template.pages:
        # 90mm×50mm尺寸定义 - 填充整个小页面
        # 90mm×50mm ≈ 255×142 点
        # 表格几乎填满整个页面，只留小边距
        table_margin = 5  # 小边距适应小尺寸
        table_x = table_margin
        table_y = table_margin
        table_width = page_width - 2 * table_margin  # 几乎填满宽度
        table_height = page_height - 2 * table_margin  # 几乎填满高度
        
        # 左列固定宽度80点，右列占剩余空间 (增加宽度避免文字被切断)
        label_width = 80
        content_width = table_width - label_width
        
        # 设置线条宽度，适应小尺寸
        c.setLineWidth(1)  # 适中线条
        
        # 绘制外框
        c.rect(table_x, table_y, table_width, table_height)
        
        # 6行布局：均分表格高度，其中Quantity占2行
        single_row_height = table_height / 6
        row_heights = [single_row_height] * 6  # 6行基础高度
        
        # 从elements中提取数据
        elements_dict = {element.role: element.content for element in page.elements}
        
        current_y = table_y + table_height
        
        # 第1行: Item
        current_y -= row_heights[0]
        c.line(table_x + label_width, current_y, table_x + label_width, current_y + row_heights[0])
        c.setFont("Helvetica-Bold", 12)  # 增大左侧标签字体
        c.drawString(table_x + 3, current_y + (row_heights[0] - 12) / 2, "Item:")
        c.setFont("Helvetica-Bold", 10)  # 增大右侧内容字体，但小于左侧
        item_content = elements_dict.get("item", "Paper Cards")
        text_width = c.stringWidth(item_content, "Helvetica-Bold", 10)
        content_x = table_x + label_width + (content_width - text_width) / 2
        c.drawString(content_x, current_y + (row_heights[0] - 10) / 2, item_content)
        
        # 第2行: Theme
        c.line(table_x, current_y, table_x + table_width, current_y)
        current_y -= row_heights[1]
        c.line(table_x + label_width, current_y, table_x + label_width, current_y + row_heights[1])
        c.setFont("Helvetica-Bold", 12)  # 增大左侧标签字体
        c.drawString(table_x + 3, current_y + (row_heights[1] - 12) / 2, "Theme:")
        c.setFont("Helvetica-Bold", 10)  # 增大右侧内容字体
        theme_content = elements_dict.get("theme", "")
        text_width = c.stringWidth(theme_content, "Helvetica-Bold", 10)
        content_x = table_x + label_width + (content_width - text_width) / 2
        c.drawString(content_x, current_y + (row_heights[1] - 10) / 2, theme_content)
        
        # 第3-4行: Quantity (占用两行高度的合并单元格)
        c.line(table_x, current_y, table_x + table_width, current_y)
        double_row_height = row_heights[2] + row_heights[3]  # 两行高度
        current_y -= double_row_height
        
        # 绘制Quantity标签的垂直边线（跨越两行）
        c.line(table_x + label_width, current_y, table_x + label_width, current_y + double_row_height)
        
        # Quantity标签居中显示在两行高度中
        c.setFont("Helvetica-Bold", 12)  # 增大左侧标签字体
        c.drawString(table_x + 3, current_y + (double_row_height - 12) / 2, "Quantity:")
        
        # 右侧内容分为上下两部分，用中线分隔
        content_area_height = double_row_height
        mid_y = current_y + content_area_height / 2
        
        # 绘制中间分隔线
        c.line(table_x + label_width, mid_y, table_x + table_width, mid_y)
        
        # 上半部分：数量
        quantity_content = elements_dict.get("quantity", "")
        c.setFont("Helvetica-Bold", 11)  # 增大右侧内容字体，稍小于左侧
        text_width = c.stringWidth(quantity_content, "Helvetica-Bold", 11)
        content_x = table_x + label_width + (content_width - text_width) / 2
        upper_center_y = mid_y + (content_area_height / 4)  # 上半部分中心
        c.drawString(content_x, upper_center_y - 5, quantity_content)
        
        # 下半部分：序列号范围
        serial_content = elements_dict.get("serial_range", "")
        c.setFont("Helvetica-Bold", 9)  # 增大序列号字体，但相对较小
        text_width = c.stringWidth(serial_content, "Helvetica-Bold", 9)
        content_x = table_x + label_width + (content_width - text_width) / 2
        lower_center_y = current_y + (content_area_height / 4)  # 下半部分中心
        c.drawString(content_x, lower_center_y - 4, serial_content)
        
        # 第5行: Carton No.
        c.line(table_x, current_y, table_x + table_width, current_y)
        current_y -= row_heights[4]
        c.line(table_x + label_width, current_y, table_x + label_width, current_y + row_heights[4])
        c.setFont("Helvetica-Bold", 12)  # 增大左侧标签字体
        c.drawString(table_x + 3, current_y + (row_heights[4] - 12) / 2, "Carton No.:")
        c.setFont("Helvetica-Bold", 10)  # 增大右侧内容字体
        carton_content = elements_dict.get("carton_no", "")
        text_width = c.stringWidth(carton_content, "Helvetica-Bold", 10)
        content_x = table_x + label_width + (content_width - text_width) / 2
        c.drawString(content_x, current_y + (row_heights[4] - 10) / 2, carton_content)
        
        # 第6行: Remark
        c.line(table_x, current_y, table_x + table_width, current_y)
        current_y -= row_heights[5]
        c.line(table_x + label_width, current_y, table_x + label_width, current_y + row_heights[5])
        c.setFont("Helvetica-Bold", 12)  # 增大左侧标签字体
        c.drawString(table_x + 3, current_y + (row_heights[5] - 12) / 2, "Remark:")
        c.setFont("Helvetica-Bold", 10)  # 增大右侧内容字体
        remark_content = elements_dict.get("remark", "")
        text_width = c.stringWidth(remark_content, "Helvetica-Bold", 10)
        content_x = table_x + label_width + (content_width - text_width) / 2
        c.drawString(content_x, current_y + (row_heights[5] - 10) / 2, remark_content)
        
        c.showPage()
    
    c.save()
    print(f"箱标PDF已生成: {output_path}")

def generate_case_pdf_from_excel(excel_path, output_path, additional_inputs=None, template_mode="regular_single"):
    """
    从Excel文件生成箱标PDF的完整流程
    
    Args:
        excel_path: Excel文件路径
        output_path: 输出PDF文件路径
        additional_inputs: 额外输入参数 (可选)
        template_mode: 箱标模板模式
    """
    
    # 1. 读取Excel数据
    excel_reader = ExcelReader(excel_path)
    excel_variables = excel_reader.extract_template_variables()
    
    print(f"读取到的Excel数据:")
    print(f"  客户编码: {excel_variables.get('customer_code', '未找到')}")
    print(f"  主题: {excel_variables.get('theme', '未找到')}")
    print(f"  开始号: {excel_variables.get('start_number', '未找到')}")
    print(f"  总张数: {excel_variables.get('total_sheets', '未找到')}")
    
    # 2. 创建箱标模板数据
    template = create_case_template_data(excel_variables, additional_inputs, template_mode)
    
    # 3. 渲染PDF
    render_case_template_to_pdf(template, output_path)

# 箱标测试配置
def get_case_test_configurations():
    """获取箱标综合测试配置"""
    import glob
    
    # 查找definition文件夹中的Excel文件
    base_path = os.path.join(os.path.dirname(__file__), '..', 'definition')
    pattern = os.path.join(base_path, '**', '*.xlsx')
    excel_files = glob.glob(pattern, recursive=True)
    
    configs = {}
    
    # 为每个Excel文件创建3种箱标模式的测试
    for excel_path in excel_files:
        filename = os.path.basename(excel_path)
        
        # 根据文件路径确定Excel类型
        if '套盒' in excel_path:
            excel_category = '套盒'
        elif '分盒' in excel_path:
            excel_category = '分盒'
        elif '常规' in filename:
            excel_category = '常规'
        else:
            excel_category = '其他'
        
        # 箱标的3种模式
        case_modes = [
            ('regular_single', '单级常规'),
            ('multi_separate', '多级分盒'),
            ('multi_set', '多级分套')
        ]
        
        for mode_key, mode_name in case_modes:
            test_key = f"{excel_category}_case_{mode_key}"
            
            configs[test_key] = {
                'name': f"{excel_category} - {mode_name}箱标",
                'excel_path': excel_path,
                'excel_name': filename,
                'excel_category': excel_category,
                'template_mode': mode_key,
                'additional_inputs': {
                    'sheets_per_box': 2850,
                    'boxes_per_small_case': 4 if mode_key != 'regular_single' else 1,
                    'small_cases_per_large_case': 2
                }
            }
    
    return configs

def run_case_test(test_key):
    """运行指定的箱标测试"""
    configs = get_case_test_configurations()
    if test_key not in configs:
        print(f"未知的箱标测试: {test_key}")
        return False
        
    config = configs[test_key]
    print(f"\\n=== 箱标测试: {config['name']} ===")
    print(f"Excel文件: {config['excel_name']}")
    print(f"Excel类型: {config['excel_category']}")
    print(f"模板模式: {config['template_mode']}")
    print(f"每盒张数: {config['additional_inputs']['sheets_per_box']}")
    print(f"每小箱盒数: {config['additional_inputs']['boxes_per_small_case']}")
    
    try:
        # 读取Excel数据以获取theme字段
        excel_reader = ExcelReader(config['excel_path'])
        excel_variables = excel_reader.extract_template_variables()
        theme = excel_variables.get('theme', '')
        
        # 生成基于主题的输出文件名
        base_filename = extract_english_filename(config['excel_path'], theme)
        output_path = f"{base_filename}_case_{config['template_mode']}.pdf"
        
        generate_case_pdf_from_excel(
            config['excel_path'], 
            output_path,
            config['additional_inputs'],
            config['template_mode']
        )
        print(f"✅ 箱标PDF生成成功: {output_path}")
        return True
    except Exception as e:
        print(f"❌ 生成失败: {e}")
        return False

def run_all_case_tests():
    """运行所有箱标测试"""
    configs = get_case_test_configurations()
    results = {}
    
    print("\\n" + "="*80)
    print("开始箱标全面测试：每个Excel文件 × 3种箱标模式")
    print("="*80)
    print(f"总测试数量: {len(configs)} 个")
    
    # 按Excel文件分组显示
    excel_groups = {}
    for test_key, config in configs.items():
        category = config['excel_category']
        if category not in excel_groups:
            excel_groups[category] = []
        excel_groups[category].append(test_key)
    
    for category, test_keys in excel_groups.items():
        print(f"\\n--- Excel文件类型: {category} ---")
        for test_key in test_keys:
            results[test_key] = run_case_test(test_key)
        
    print("\\n" + "="*80) 
    print("箱标测试结果汇总:")
    print("="*80)
    
    # 按Excel类型分组显示结果
    for category, test_keys in excel_groups.items():
        print(f"\\n{category}:")
        for test_key in test_keys:
            config = configs[test_key]
            success = results[test_key]
            status = "✅ 成功" if success else "❌ 失败"
            mode_name = config['name'].split(' - ')[1] if ' - ' in config['name'] else config['name']
            print(f"  {mode_name:15} - {status}")
    
    # 统计结果
    total_tests = len(results)
    successful_tests = sum(results.values())
    print(f"\\n总体结果: {successful_tests}/{total_tests} 成功")
    
    return results

# 箱标渲染程序入口
if __name__ == "__main__":
    # 运行所有箱标测试
    run_all_case_tests()
    
    # 也可以单独运行某个测试：
    # run_case_test("套盒_case_regular_single")  # 仅测试套盒的单级常规箱标
    # run_case_test("分盒_case_multi_separate")  # 仅测试分盒的多级分盒箱标
    # run_case_test("常规_case_multi_set")       # 仅测试常规的多级分套箱标