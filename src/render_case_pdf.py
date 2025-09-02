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
        cards_per_box = additional_inputs.get('sheets_per_box', 2850)
        boxes_per_small_case = additional_inputs.get('boxes_per_small_case', 1)
        small_cases_per_large_case = additional_inputs.get('small_cases_per_large_case', 2)
    else:
        cards_per_box = 2850
        boxes_per_small_case = 1  
        small_cases_per_large_case = 2
    
    # 计算箱标数量和编号
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
                CaseElement(role="carton_no", content=f"{start_small_case + 1}/{end_small_case + 1}"),
                CaseElement(role="remark", content=customer_code)
            ]))
            
    elif template_mode == "multi_separate":
        # =============================================
        # 多级分盒模式：小箱标(单多级盒) + 大箱标(多多级盒)
        # =============================================
        
        # 小箱标: 单个多级盒子 (3500PCS, LGM01001-01-LGM01001-01, 1-1)
        for case_idx in range(total_small_cases):
            for sub_box in range(boxes_per_small_case):
                box_num = start_num + case_idx
                sub_serial = f"{serial_prefix}{box_num:05d}-{sub_box + 1:02d}"
                
                pages.append(CasePage(index=case_idx*boxes_per_small_case + sub_box + 1, elements=[
                    CaseElement(role="item", content="Paper Cards"),
                    CaseElement(role="theme", content=f"{theme} (chip)"),
                    CaseElement(role="quantity", content=f"{cards_per_box}PCS"),
                    CaseElement(role="serial_range", content=f"{sub_serial}-{sub_serial}"),
                    CaseElement(role="carton_no", content=f"{case_idx + 1}-{sub_box + 1}"),
                    CaseElement(role="remark", content=customer_code)
                ]))
        
        # 大箱标: 多个多级盒子范围 (7000PCS, LGM01029-01-LGM01030-02, 29-30)
        for large_case_idx in range(total_large_cases):
            start_case = large_case_idx * small_cases_per_large_case
            end_case = min(start_case + small_cases_per_large_case - 1, total_small_cases - 1)
            
            start_box_num = start_num + start_case
            end_box_num = start_num + end_case
            
            start_serial = f"{serial_prefix}{start_box_num:05d}-01"
            end_serial = f"{serial_prefix}{end_box_num:05d}-{boxes_per_small_case:02d}"
            
            large_quantity = (end_case - start_case + 1) * boxes_per_small_case * cards_per_box
            
            pages.append(CasePage(index=len(pages) + 1, elements=[
                CaseElement(role="item", content="Paper Cards"),
                CaseElement(role="theme", content=f"{theme} (chip)"),
                CaseElement(role="quantity", content=f"{large_quantity}PCS"),
                CaseElement(role="serial_range", content=f"{start_serial}-{end_serial}"),
                CaseElement(role="carton_no", content=f"{start_case + 1}-{end_case + 1}"),
                CaseElement(role="remark", content=customer_code)
            ]))
            
    elif template_mode == "multi_set":
        # =============================================
        # 多级分套模式：小箱标(套内多级) + 大箱标(跨套多级)
        # =============================================
        
        # 小箱标: 套内多级范围 (3780PCS, JAW01001-01-JAW01001-06, 01)
        for case_idx in range(total_small_cases):
            box_num = start_num + case_idx
            start_serial = f"{serial_prefix}{box_num:05d}-01"
            end_serial = f"{serial_prefix}{box_num:05d}-{boxes_per_small_case:02d}"
            
            small_quantity = boxes_per_small_case * cards_per_box
            
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
            
            start_box_num = start_num + start_case
            end_box_num = start_num + end_case
            
            start_serial = f"{serial_prefix}{start_box_num:05d}-01"
            end_serial = f"{serial_prefix}{end_box_num:05d}-{boxes_per_small_case:02d}"
            
            large_quantity = (end_case - start_case + 1) * boxes_per_small_case * cards_per_box
            
            pages.append(CasePage(index=large_case_idx*2 + 2, elements=[
                CaseElement(role="item", content="Paper Cards"),
                CaseElement(role="theme", content=theme),
                CaseElement(role="quantity", content=f"{large_quantity}PCS"),
                CaseElement(role="serial_range", content=f"{start_serial}-{end_serial}"),
                CaseElement(role="carton_no", content=f"{start_case + start_num}-{end_case + start_num}"),
                CaseElement(role="remark", content=customer_code)
            ]))
    
    # 箱标使用横向A4布局
    return CaseTemplate(
        width_pt=landscape(A4)[0],
        height_pt=landscape(A4)[1], 
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
    c = canvas.Canvas(output_path, pagesize=landscape(A4))
    page_width, page_height = landscape(A4)
    
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
        # 完全按照目标样式的绝对尺寸定义 - 填充整个纸张
        # A4横向尺寸：842x595 点
        # 表格几乎填满整个纸张，只留极小边距
        table_margin = 20  # 极小边距
        table_x = table_margin
        table_y = table_margin
        table_width = page_width - 2 * table_margin  # 几乎填满宽度
        table_height = page_height - 2 * table_margin  # 几乎填满高度
        
        # 左列固定宽度200点，右列占剩余空间
        label_width = 200
        content_width = table_width - label_width
        
        # 设置粗线条
        c.setLineWidth(3)  # 更粗的边框
        
        # 绘制外框
        c.rect(table_x, table_y, table_width, table_height)
        
        # 每行固定高度 - 平均分配
        row_height = table_height / 5
        row_heights = [row_height] * 5  # 5行相等高度
        
        # 从elements中提取数据
        elements_dict = {element.role: element.content for element in page.elements}
        
        current_y = table_y + table_height
        
        # 第1行: Item
        current_y -= row_heights[0]
        # 绘制列分隔线 (粗线)
        c.line(table_x + label_width, current_y, table_x + label_width, current_y + row_heights[0])
        # 标签 - 绝对大小字体，左对齐有边距
        c.setFont("MicrosoftYaHei", 40)  # 绝对大字体标签
        c.drawString(table_x + 30, current_y + (row_heights[0] - 40) / 2, "Item:")
        # 内容 - 绝对大字体，严格居中
        c.setFont("MicrosoftYaHei", 45)  # 绝对大内容字体
        item_content = elements_dict.get("item", "Paper Cards")
        text_width = c.stringWidth(item_content, "MicrosoftYaHei", 45)
        content_x = table_x + label_width + (content_width - text_width) / 2
        c.drawString(content_x, current_y + (row_heights[0] - 45) / 2, item_content)
        
        # 第2行: Theme
        c.line(table_x, current_y, table_x + table_width, current_y)  # 行分隔线 (粗线)
        current_y -= row_heights[1]
        c.line(table_x + label_width, current_y, table_x + label_width, current_y + row_heights[1])
        c.setFont("MicrosoftYaHei", 40)
        c.drawString(table_x + 30, current_y + (row_heights[1] - 40) / 2, "Theme:")
        c.setFont("MicrosoftYaHei", 45)
        theme_content = elements_dict.get("theme", "")
        text_width = c.stringWidth(theme_content, "MicrosoftYaHei", 45)
        content_x = table_x + label_width + (content_width - text_width) / 2
        c.drawString(content_x, current_y + (row_heights[1] - 45) / 2, theme_content)
        
        # 第3行: Quantity (均匀行高，包含双层内容)
        c.line(table_x, current_y, table_x + table_width, current_y)
        current_y -= row_heights[2]
        c.line(table_x + label_width, current_y, table_x + label_width, current_y + row_heights[2])
        c.setFont("MicrosoftYaHei", 40)
        c.drawString(table_x + 30, current_y + (row_heights[2] - 40) / 2, "Quantity:")
        
        # 上半部分：数量 - 超大字体，居中
        quantity_content = elements_dict.get("quantity", "")
        c.setFont("MicrosoftYaHei", 42)  # 数量绝对大字体
        text_width = c.stringWidth(quantity_content, "MicrosoftYaHei", 42)
        content_x = table_x + label_width + (content_width - text_width) / 2
        # 调整位置，使双层内容在行内均匀分布
        c.drawString(content_x, current_y + row_heights[2] * 0.7, quantity_content)
        
        # 下半部分：序列号范围 - 大字体，居中
        serial_content = elements_dict.get("serial_range", "")
        c.setFont("MicrosoftYaHei", 35)  # 序列号绝对大字体
        text_width = c.stringWidth(serial_content, "MicrosoftYaHei", 35)
        content_x = table_x + label_width + (content_width - text_width) / 2
        c.drawString(content_x, current_y + row_heights[2] * 0.25, serial_content)
        
        # 第4行: Carton No.
        c.line(table_x, current_y, table_x + table_width, current_y)
        current_y -= row_heights[3]
        c.line(table_x + label_width, current_y, table_x + label_width, current_y + row_heights[3])
        c.setFont("MicrosoftYaHei", 40)
        c.drawString(table_x + 30, current_y + (row_heights[3] - 40) / 2, "Carton No.:")
        c.setFont("MicrosoftYaHei", 45)
        carton_content = elements_dict.get("carton_no", "")
        text_width = c.stringWidth(carton_content, "MicrosoftYaHei", 45)
        content_x = table_x + label_width + (content_width - text_width) / 2
        c.drawString(content_x, current_y + (row_heights[3] - 45) / 2, carton_content)
        
        # 第5行: Remark
        c.line(table_x, current_y, table_x + table_width, current_y)
        current_y -= row_heights[4]
        c.line(table_x + label_width, current_y, table_x + label_width, current_y + row_heights[4])
        c.setFont("MicrosoftYaHei", 40)
        c.drawString(table_x + 30, current_y + (row_heights[4] - 40) / 2, "Remark:")
        c.setFont("MicrosoftYaHei", 45)
        remark_content = elements_dict.get("remark", "")
        text_width = c.stringWidth(remark_content, "MicrosoftYaHei", 45)
        content_x = table_x + label_width + (content_width - text_width) / 2
        c.drawString(content_x, current_y + (row_heights[4] - 45) / 2, remark_content)
        
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