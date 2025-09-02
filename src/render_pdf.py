import os
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

class Template:
    def __init__(self, width_pt, height_pt, style_refs, pages):
        self.width_pt = width_pt
        self.height_pt = height_pt
        self.pages = pages

class Page:
    def __init__(self, index=None, index_range=None, elements=[]):
        self.index = index
        self.index_range = index_range
        self.elements = elements

class Element:
    def __init__(self, role, content=None, content_template=None, seq=None):
        self.role = role
        self.content = content
        self.content_template = content_template
        self.seq = seq

class Sequence:
    def __init__(self, start, end, pad):
        self.start = start
        self.end = end
        self.pad = pad

class Style:
    def __init__(self, font_family, font_weight, font_size_pt, x_align, y_top_pt):
        self.font_family = font_family
        self.font_weight = font_weight
        self.font_size_pt = font_size_pt
        self.x_align = x_align
        self.y_top_pt = y_top_pt

def render_pdf(template, style_set, output_filename):
    c = canvas.Canvas(output_filename, pagesize=landscape(A4))
    
    for page in template.pages:
        if page.index is not None:
            render_page(c, page, style_set)
        elif page.index_range is not None:
            for index in range(page.index_range[0], page.index_range[1] + 1):
                render_page(c, page, style_set, index)
    
    c.save()

def render_page(c, page, style_set, index=None):
    page_height = c._pagesize[1]
    # 先绘制内容，再绘制序列号
    for element in sorted(page.elements, key=lambda e: e.role, reverse=True):
        style = style_set[element.role]
        content = element.content
        if element.seq and element.content_template:
            seq_num = str(index).zfill(element.seq.pad)
            # 支持不同的序列号占位符
            content = element.content_template.replace("{seq3}", seq_num)
            content = content.replace("{seq5}", seq_num)

        # 设置字体
        c.setFont(style.font_family, style.font_size_pt)

        text_width = c.stringWidth(content, style.font_family, style.font_size_pt)
        x_position = calculate_x_position(style.x_align, text_width, c._pagesize[0])
        # 计算垂直居中的 y_position
        y_position = (page_height + style.y_top_pt) / 2
        c.drawString(x_position, y_position, content)
        
    c.showPage()

def calculate_x_position(x_align, text_width, page_width):
    if x_align == "center":
        return (page_width - text_width) / 2
    elif x_align == "left":
        return 0
    elif x_align == "right":
        return page_width - text_width

def create_template_data(excel_variables, additional_inputs=None):
    """
    根据Excel数据和额外输入值创建template_data
    
    Args:
        excel_variables: 从Excel提取的变量 {
            'customer_code': str,  # A4: 客户名称编码
            'theme': str,          # B4: 主题  
            'start_number': str,   # B11: 开始号
            'total_sheets': int    # F4: 总张数
        }
        additional_inputs: 额外输入值 (可选) {
            'sheets_per_box': int,      # 分盒张数
            'sheets_per_set': int,      # 分套张数  
            'boxes_per_small_case': int, # 小箱内的盒数
            'small_cases_per_large_case': int # 大箱内的小箱数
        }
    
    Returns:
        Template对象
    """
    
    # 分盒模式下的两级编号计算
    total_cards = excel_variables.get('total_sheets', 1)  # F3的总张数(实际是卡片数)
    
    if additional_inputs and 'sheets_per_box' in additional_inputs:
        # 获取分盒参数
        cards_per_box = additional_inputs.get('sheets_per_box', 1)  # 每盒卡片数
        boxes_per_small_case = additional_inputs.get('boxes_per_small_case', 1)  # 每小箱盒数
        
        # 计算总盒数和总小箱数 - 使用向上取整
        import math
        total_boxes = math.ceil(total_cards / cards_per_box) if cards_per_box > 0 else 1
        total_small_cases = math.ceil(total_boxes / boxes_per_small_case) if boxes_per_small_case > 0 else 1
        
        # 存储分盒信息供后续使用
        box_info = {
            'total_boxes': total_boxes,
            'total_small_cases': total_small_cases, 
            'boxes_per_small_case': boxes_per_small_case
        }
    else:
        # 常规模式：直接使用总数
        total_boxes = total_cards
        box_info = {'total_boxes': total_boxes}
    
    # 提取开始号中的前缀和数字部分
    start_number_str = excel_variables.get('start_number', '1')
    try:
        import re
        # 提取字母前缀和数字部分
        match = re.match(r'([A-Za-z]+)(\d+)', start_number_str)
        if match:
            serial_prefix = match.group(1)  # 字母前缀 (如 LAN)
            start_num = int(match.group(2))  # 数字部分 (如 01001)
        else:
            # 如果没有字母前缀，尝试提取纯数字
            numbers = re.findall(r'\d+', start_number_str)
            serial_prefix = excel_variables.get('customer_code', '')  # 备用使用客户编码
            start_num = int(numbers[-1]) if numbers else 1
    except:
        serial_prefix = excel_variables.get('customer_code', '')
        start_num = 1
    
    customer_code = excel_variables.get('customer_code', '')
    theme = excel_variables.get('theme', '')
    
    # 根据是否是分盒模式生成不同的页面
    pages = [
        # 首页显示客户信息和主题
        Page(index=0, elements=[
            Element(role="customer_info", content=customer_code),
            Element(role="theme_info", content=theme)
        ])
    ]
    
    if 'boxes_per_small_case' in box_info and box_info.get('boxes_per_small_case', 1) > 1:
        # 分盒模式：生成两级编号页面
        page_index = 1
        small_case_num = start_num
        
        boxes_generated = 0
        
        for case_idx in range(box_info['total_small_cases']):
            case_serial = f"{serial_prefix}{small_case_num:05d}"  # 小箱编号 LGM01001
            
            # 计算这个小箱应该有多少个盒子
            remaining_boxes = box_info['total_boxes'] - boxes_generated
            boxes_in_this_case = min(box_info['boxes_per_small_case'], remaining_boxes)
            
            for box_idx in range(boxes_in_this_case):
                box_serial = f"{case_serial}-{box_idx + 1:02d}"  # 盒子编号 LGM01001-01
                
                pages.append(Page(index=page_index, elements=[
                    Element(role="theme_info", content=theme),
                    Element(role="serial_number", content=box_serial)
                ]))
                
                page_index += 1
                boxes_generated += 1
            
            small_case_num += 1
            
            # 如果已经生成了所有盒子，停止
            if boxes_generated >= box_info['total_boxes']:
                break
    else:
        # 常规模式：单级编号
        end_num = start_num + box_info['total_boxes'] - 1
        pages.append(Page(index_range=[start_num, end_num], elements=[
            Element(role="theme_info", content=theme),
            Element(
                role="serial_number", 
                content_template=f"{serial_prefix}{{seq5}}", 
                seq=Sequence(start=start_num, end=end_num, pad=5)
            )
        ]))
    
    return Template(
        width_pt=landscape(A4)[0],
        height_pt=landscape(A4)[1], 
        style_refs=["customer_info", "theme_info", "serial_number"],
        pages=pages
    )

def calculate_total_sheets(additional_inputs):
    """根据包装参数计算总张数(备用方法)"""
    sheets_per_box = additional_inputs.get('sheets_per_box', 1)
    boxes_per_small_case = additional_inputs.get('boxes_per_small_case', 1)
    small_cases_per_large_case = additional_inputs.get('small_cases_per_large_case', 1)
    
    return sheets_per_box * boxes_per_small_case * small_cases_per_large_case

def generate_pdf_from_excel(excel_path, output_path, additional_inputs=None):
    """
    从Excel文件生成PDF的完整流程
    
    Args:
        excel_path: Excel文件路径
        output_path: 输出PDF文件路径
        additional_inputs: 额外输入参数 (可选)
    """
    
    # 1. 读取Excel数据
    excel_reader = ExcelReader(excel_path)
    excel_variables = excel_reader.extract_template_variables()
    
    print(f"读取到的Excel数据:")
    print(f"  客户编码: {excel_variables.get('customer_code', '未找到')}")
    print(f"  主题: {excel_variables.get('theme', '未找到')}")
    print(f"  开始号: {excel_variables.get('start_number', '未找到')}")
    print(f"  总张数: {excel_variables.get('total_sheets', '未找到')}")
    
    # 2. 创建模板数据
    template = create_template_data(excel_variables, additional_inputs)
    
    # 3. 创建样式集
    style_set = {
        "customer_info": Style(
            font_family="MicrosoftYaHei", 
            font_weight="bold", 
            font_size_pt=20, 
            x_align="center", 
            y_top_pt=300
        ),
        "theme_info": Style(
            font_family="MicrosoftYaHei", 
            font_weight="bold", 
            font_size_pt=24, 
            x_align="center", 
            y_top_pt=250
        ),
        "serial_number": Style(
            font_family="Helvetica-Bold", 
            font_weight="bold", 
            font_size_pt=18, 
            x_align="center", 
            y_top_pt=200
        )
    }
    
    # 4. 生成PDF
    render_pdf(template, style_set, output_path)
    print(f"PDF已生成: {output_path}")
    
    return output_path

# 示例数据
template_data = Template(
    width_pt=landscape(A4)[0],
    height_pt=landscape(A4)[1],
    style_refs=["title_cn", "title_en", "serial"],
    pages=[
        Page(index=0, elements=[
            Element(role="title_cn", content="德克斯的助手")
        ]),
        Page(index_range=[1, 100], elements=[
            Element(role="title_en", content="DEX'S SIDEKICK"),
            Element(role="serial", content_template="DSK{seq3}", seq=Sequence(start=1, end=100, pad=3))
        ])
    ]
)

style_set_data = {
    "title_cn": Style(font_family="MicrosoftYaHei", font_weight="bold", font_size_pt=24, x_align="center", y_top_pt=250),
    "title_en": Style(font_family="Helvetica-Bold", font_weight="bold", font_size_pt=24, x_align="center", y_top_pt=200),
    "serial": Style(font_family="Helvetica-Bold", font_weight="bold", font_size_pt=24, x_align="center", y_top_pt=150)
}

# 使用Excel数据生成PDF的示例
if __name__ == "__main__":
    # 示例1: 使用Excel文件生成PDF
    try:
        # 直接指定Excel文件路径
        # excel_path = "/Users/qyx/Desktop/task/data-to-pdfprint/definition/常规/模板1/常规-LADIES NIGHT IN 女士夜..xlsx"
        excel_path = "/Users/qyx/Desktop/task/data-to-pdfprint/definition/套盒/套盒-JAWS-大白鲨 .xlsx"
        # 检查文件是否存在
        if os.path.exists(excel_path):
            excel_files = [excel_path]
        else:
            # 如果指定文件不存在，搜索目录中的Excel文件
            input_dirs = ["definition", "definition/常规/模板1"]
            excel_files = []
            
            for input_dir in input_dirs:
                if os.path.exists(input_dir):
                    for root, dirs, files in os.walk(input_dir):
                        for file in files:
                            if file.endswith(('.xlsx', '.xls')):
                                excel_files.append(os.path.join(root, file))
                    break  # 找到第一个存在的目录就停止
        
        if excel_files:
            excel_path = excel_files[0]  # 使用第一个找到的Excel文件
            output_path = "output_from_excel1.pdf"
            
            # 可选的额外输入参数
            additional_inputs = {
                'sheets_per_box': 3780,  # 每盒3780张卡片
                'boxes_per_small_case': 2,  # 每小箱2盒（这样20盒会分成5个小箱）
                'small_cases_per_large_case': 5
            }
            
            generate_pdf_from_excel(excel_path, output_path, additional_inputs)
        else:
            print("未找到Excel文件，使用示例数据生成PDF")
            render_pdf(template_data, style_set_data, "output4.pdf")
            
    except Exception as e:
        print(f"生成PDF时出错: {e}")
        print("使用示例数据生成PDF")
        render_pdf(template_data, style_set_data, "output4.pdf")
