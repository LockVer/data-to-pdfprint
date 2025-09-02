import os
import re
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
        english_only = "output"
    
    return english_only

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

def create_template_data(excel_variables, additional_inputs=None, template_mode="two_level"):
    """
    根据Excel数据和额外输入值创建template_data
    
    Args:
        excel_variables: 从Excel提取的变量 {
            'customer_code': str,  # A3: 客户名称编码
            'theme': str,          # B3: 主题  
            'start_number': str,   # B10: 开始号
            'total_sheets': int    # F3: 总张数
        }
        additional_inputs: 额外输入值 (可选) {
            'sheets_per_box': int,      # 分盒张数
            'sheets_per_set': int,      # 分套张数  
            'boxes_per_small_case': int, # 小箱内的盒数
            'small_cases_per_large_case': int # 大箱内的小箱数
        }
        template_mode: 模板模式 {
            'two_level': 两级编号模式 (LGM01001-01)
            'game_info': 游戏信息模式 (Game title + Ticket count + Serial)
        }
    
    Returns:
        Template对象
    """
    
    # ===========================================
    # 计算盒子数量和编号级数
    # ===========================================
    total_cards = excel_variables.get('total_sheets', 1)  # F3的总张数(实际是卡片数)
    
    if additional_inputs and 'sheets_per_box' in additional_inputs:
        # -------------------------------------------
        # 分盒模式：根据输入参数计算盒数和小箱数
        # -------------------------------------------
        cards_per_box = additional_inputs.get('sheets_per_box', 1)      # 每盒卡片数 (如: 3780)
        boxes_per_small_case = additional_inputs.get('boxes_per_small_case', 1)  # 每小箱盒数 (如: 4)
        
        # 使用向上取整确保所有卡片都有对应的盒子
        import math
        total_boxes = math.ceil(total_cards / cards_per_box) if cards_per_box > 0 else 1      # 总盒数
        total_small_cases = math.ceil(total_boxes / boxes_per_small_case) if boxes_per_small_case > 0 else 1  # 总小箱数
        
        # 判断序列号级数：
        # boxes_per_small_case = 1: 单级编号 (JAW01001, JAW01002...)  
        # boxes_per_small_case > 1: 多级编号 (JAW01001-01, JAW01001-02...)
        box_info = {
            'total_boxes': total_boxes,
            'total_small_cases': total_small_cases, 
            'boxes_per_small_case': boxes_per_small_case
        }
    else:
        # -------------------------------------------
        # 常规模式：直接使用Excel总数，默认单级编号
        # -------------------------------------------
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
    
    # ===========================================
    # 根据模板模式生成不同的页面内容
    # 支持4种标签组合类型：
    # 1. 游戏+多级：Game title + Ticket count + JAW01001-01
    # 2. 游戏+单级：Game title + Ticket count + JAW01001
    # 3. 常规+多级：首页(客户+主题) + 其他页(主题+JAW01001-01)
    # 4. 常规+单级：首页(客户+主题) + 其他页(主题+JAW01001)
    # ===========================================

    if template_mode == "game_info":
        # ===================================================
        # 游戏信息模式：每页显示Game title + Ticket count + Serial
        # ===================================================
        pages = []
        # TODO 这里首页显示Game title的中文字段值 + Ticket count空 + Serial空
        if 'boxes_per_small_case' in box_info and box_info.get('boxes_per_small_case', 1) > 1:
            # =====================================================
            # 【游戏+多级】分盒模式：每页显示游戏信息+多级序列号
            # 序列号格式：JAW01001-01, JAW01001-02, JAW01002-01...
            # =====================================================
            page_index = 1
            small_case_num = start_num
            boxes_generated = 0
            
            for case_idx in range(box_info['total_small_cases']):
                case_serial = f"{serial_prefix}{small_case_num:05d}"  # 小箱编号 JAW01001
                
                remaining_boxes = box_info['total_boxes'] - boxes_generated
                boxes_in_this_case = min(box_info['boxes_per_small_case'], remaining_boxes)
                
                for box_idx in range(boxes_in_this_case):
                    box_serial = f"{case_serial}-{box_idx + 1:02d}"  # 多级序列号 JAW01001-01
                    
                    # 获取每盒卡片数
                    cards_per_box = additional_inputs.get('sheets_per_box', total_cards) if additional_inputs else total_cards
                    
                    pages.append(Page(index=page_index, elements=[
                        Element(role="game_title", content=f"Game title: {theme}"),        # Game title: JAWS
                        Element(role="ticket_count", content=f"Ticket count: {cards_per_box}"), # Ticket count: 3780
                        Element(role="serial_number", content=f"Serial: {box_serial}")     # Serial: JAW01001-01
                    ]))
                    
                    page_index += 1
                    boxes_generated += 1
                
                small_case_num += 1
                if boxes_generated >= box_info['total_boxes']:
                    break
        else:
            # =====================================================
            # 【游戏+单级】常规模式：每页显示游戏信息+单级序列号
            # 序列号格式：JAW01001, JAW01002, JAW01003...
            # =====================================================
            cards_per_box = additional_inputs.get('sheets_per_box', total_cards) if additional_inputs else total_cards
            
            for i in range(box_info['total_boxes']):
                serial_num = start_num + i
                serial = f"{serial_prefix}{serial_num:05d}"  # 单级序列号 JAW01001
                
                pages.append(Page(index=i+1, elements=[
                    Element(role="game_title", content=f"Game title: {theme}"),        # Game title: JAWS
                    Element(role="ticket_count", content=f"Ticket count: {cards_per_box}"), # Ticket count: 3780
                    Element(role="serial_number", content=f"Serial: {serial}")         # Serial: JAW01001
                ]))
    else:
        # ===================================================  
        # 常规信息模式：首页显示客户+主题，其他页显示主题+序列号
        # ===================================================
        pages = [
            # 首页显示客户信息和主题（所有常规模式共用）
            Page(index=0, elements=[
                Element(role="customer_info", content=customer_code),  # 14KH0095
                Element(role="theme_info", content=theme)              # JAWS
            ])
        ]
        
        if 'boxes_per_small_case' in box_info and box_info.get('boxes_per_small_case', 1) > 1:
            # =====================================================
            # 【常规+多级】分盒模式：首页(客户+主题) + 其他页(主题+多级序列号)
            # 序列号格式：JAW01001-01, JAW01001-02, JAW01002-01...
            # =====================================================
            page_index = 1
            small_case_num = start_num
            boxes_generated = 0
            
            for case_idx in range(box_info['total_small_cases']):
                case_serial = f"{serial_prefix}{small_case_num:05d}"  # 小箱编号 JAW01001
                
                # 计算这个小箱应该有多少个盒子
                remaining_boxes = box_info['total_boxes'] - boxes_generated
                boxes_in_this_case = min(box_info['boxes_per_small_case'], remaining_boxes)
                
                for box_idx in range(boxes_in_this_case):
                    box_serial = f"{case_serial}-{box_idx + 1:02d}"  # 多级序列号 JAW01001-01
                    
                    pages.append(Page(index=page_index, elements=[
                        Element(role="theme_info", content=theme),          # JAWS
                        Element(role="serial_number", content=box_serial)   # JAW01001-01
                    ]))
                    
                    page_index += 1
                    boxes_generated += 1
                
                small_case_num += 1
                
                # 如果已经生成了所有盒子，停止
                if boxes_generated >= box_info['total_boxes']:
                    break
        else:
            # =====================================================
            # 【常规+单级】常规模式：首页(客户+主题) + 其他页(主题+单级序列号)
            # 序列号格式：JAW01001, JAW01002, JAW01003...
            # 使用index_range和content_template实现批量生成
            # =====================================================
            end_num = start_num + box_info['total_boxes'] - 1
            pages.append(Page(index_range=[start_num, end_num], elements=[
                Element(role="theme_info", content=theme),                    # JAWS
                Element(
                    role="serial_number", 
                    content_template=f"{serial_prefix}{{seq5}}",              # JAW{seq5} -> JAW01001
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

def generate_pdf_from_excel(excel_path, output_path, additional_inputs=None, template_mode="two_level"):
    """
    从Excel文件生成PDF的完整流程
    
    Args:
        excel_path: Excel文件路径
        output_path: 输出PDF文件路径
        additional_inputs: 额外输入参数 (可选)
        template_mode: 模板模式 ("two_level"或"game_info")
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
    template = create_template_data(excel_variables, additional_inputs, template_mode)
    
    # 3. 根据模板模式创建样式集
    if template_mode == "game_info":
        style_set = {
            "game_title": Style(
                font_family="MicrosoftYaHei", 
                font_weight="bold", 
                font_size_pt=24, 
                x_align="left", 
                y_top_pt=350
            ),
            "ticket_count": Style(
                font_family="MicrosoftYaHei", 
                font_weight="bold", 
                font_size_pt=22, 
                x_align="left", 
                y_top_pt=200
            ),
            "serial_number": Style(
                font_family="Helvetica-Bold", 
                font_weight="bold", 
                font_size_pt=20, 
                x_align="left", 
                y_top_pt=50
            )
        }
    else:
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

# ===========================================
# 4种模式组合的测试数据配置
# ===========================================

def get_comprehensive_test_configurations():
    """获取全面的测试配置：每个Excel文件 × 4种模式组合"""
    import glob
    import os
    base_path = "/Users/qyx/Desktop/task/data-to-pdfprint/definition"
    
    # 动态查找所有Excel文件
    excel_files = []
    for root, dirs, files in os.walk(base_path):
        for file in files:
            if file.endswith('.xlsx'):
                excel_files.append({
                    'path': os.path.join(root, file),
                    'name': file,
                    'category': os.path.basename(root)
                })
    
    # 定义4种模式组合的配置
    mode_configs = {
        "regular_single": {
            "name": "常规+单级",
            "template_mode": "two_level",
            "additional_inputs": {
                'sheets_per_box': 2850,
                'boxes_per_small_case': 1,  # 单级
                'small_cases_per_large_case': 5
            }
        },
        "regular_multi": {
            "name": "常规+多级", 
            "template_mode": "two_level",
            "additional_inputs": {
                'sheets_per_box': 2850,
                'boxes_per_small_case': 4,  # 多级
                'small_cases_per_large_case': 5
            }
        },
        "game_single": {
            "name": "游戏+单级",
            "template_mode": "game_info",
            "additional_inputs": {
                'sheets_per_box': 2850,
                'boxes_per_small_case': 1,  # 单级
                'small_cases_per_large_case': 5
            }
        },
        "game_multi": {
            "name": "游戏+多级",
            "template_mode": "game_info", 
            "additional_inputs": {
                'sheets_per_box': 2850,
                'boxes_per_small_case': 4,  # 多级
                'small_cases_per_large_case': 5
            }
        }
    }
    
    # 生成完整的测试矩阵：每个Excel文件 × 4种模式
    test_matrix = {}
    for excel_file in excel_files:
        for mode_key, mode_config in mode_configs.items():
            test_key = f"{excel_file['category']}_{mode_key}"
            test_matrix[test_key] = {
                "name": f"{excel_file['category']} - {mode_config['name']}",
                "excel_path": excel_file['path'],
                "excel_name": excel_file['name'],
                "excel_category": excel_file['category'],
                "template_mode": mode_config['template_mode'],
                "additional_inputs": mode_config['additional_inputs']
            }
    
    return test_matrix

def run_comprehensive_test(test_key):
    """运行指定的综合测试"""
    configs = get_comprehensive_test_configurations()
    if test_key not in configs:
        print(f"未知的测试: {test_key}")
        return False
        
    config = configs[test_key]
    print(f"\\n=== 测试: {config['name']} ===")
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
        
        # 生成基于主题或英文文件名的输出文件名
        base_filename = extract_english_filename(config['excel_path'], theme)
        
        # 根据模式类型添加后缀
        mode_suffix = ""
        if config['template_mode'] == "two_level":
            if config['additional_inputs']['boxes_per_small_case'] > 1:
                mode_suffix = "_regular_multi"
            else:
                mode_suffix = "_regular_single"
        else:  # game_info
            if config['additional_inputs']['boxes_per_small_case'] > 1:
                mode_suffix = "_game_multi"
            else:
                mode_suffix = "_game_single"
        
        output_path = f"{base_filename}{mode_suffix}.pdf"
        
        generate_pdf_from_excel(
            config['excel_path'], 
            output_path,
            config['additional_inputs'],
            config['template_mode']
        )
        print(f"✅ PDF生成成功: {output_path}")
        return True
    except Exception as e:
        print(f"❌ 生成失败: {e}")
        return False

def run_all_comprehensive_tests():
    """运行所有Excel文件的4种模式测试"""
    configs = get_comprehensive_test_configurations()
    results = {}
    
    print("\\n" + "="*80)
    print("开始全面测试：每个Excel文件 × 4种标签组合类型")
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
            results[test_key] = run_comprehensive_test(test_key)
        
    print("\\n" + "="*80) 
    print("测试结果汇总:")
    print("="*80)
    
    # 按Excel类型分组显示结果
    for category, test_keys in excel_groups.items():
        print(f"\\n{category}:")
        for test_key in test_keys:
            config = configs[test_key]
            success = results[test_key]
            status = "✅ 成功" if success else "❌ 失败"
            mode_name = config['name'].split(' - ')[1] if ' - ' in config['name'] else config['name']
            print(f"  {mode_name:12} - {status}")
    
    # 统计结果
    total_tests = len(results)
    successful_tests = sum(results.values())
    print(f"\\n总体结果: {successful_tests}/{total_tests} 成功")
    
    return results

def run_excel_file_tests(excel_category):
    """运行指定Excel文件的所有4种模式测试"""
    configs = get_comprehensive_test_configurations()
    
    # 找到指定类型的所有测试
    target_tests = {k: v for k, v in configs.items() if v['excel_category'] == excel_category}
    
    if not target_tests:
        print(f"未找到类型为 '{excel_category}' 的Excel文件")
        return {}
    
    print(f"\\n=== 测试Excel文件类型: {excel_category} ===")
    print(f"测试数量: {len(target_tests)} 个模式")
    
    results = {}
    for test_key in target_tests:
        results[test_key] = run_comprehensive_test(test_key)
    
    return results

# 使用Excel数据生成PDF的示例
if __name__ == "__main__":
    # 运行全面测试：每个Excel文件 × 4种标签组合类型
    run_all_comprehensive_tests()
    
    # 也可以单独运行某个Excel文件的所有模式：
    # run_excel_file_tests("套盒")  # 仅测试套盒Excel文件的4种模式
    # run_excel_file_tests("分盒")  # 仅测试分盒Excel文件的4种模式
    # run_excel_file_tests("常规")  # 仅测试常规Excel文件的4种模式
