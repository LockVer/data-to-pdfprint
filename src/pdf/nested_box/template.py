"""
Nested Box Template - Multi-level PDF generation with Excel serial number ranges
"""
import math
import sys
import os
from pathlib import Path
from typing import Dict, Any
from reportlab.pdfgen import canvas
from reportlab.lib.colors import CMYKColor
from reportlab.lib.units import mm

# 导入基础工具类
from src.utils.pdf_base import PDFBaseUtils
from src.utils.font_manager import font_manager
from src.utils.text_processor import text_processor
from src.utils.excel_data_extractor import ExcelDataExtractor

# 导入套盒模板专属数据处理器和渲染器
current_dir = os.path.dirname(__file__)
sys.path.insert(0, current_dir)

# 使用importlib导入以避免命名冲突
import importlib.util

# 导入data_processor
dp_spec = importlib.util.spec_from_file_location("nested_box_data_processor", os.path.join(current_dir, "data_processor.py"))
dp_module = importlib.util.module_from_spec(dp_spec)
dp_spec.loader.exec_module(dp_module)
nested_box_data_processor = dp_module.nested_box_data_processor

# 导入renderer
renderer_spec = importlib.util.spec_from_file_location("nested_box_renderer", os.path.join(current_dir, "renderer.py"))
renderer_module = importlib.util.module_from_spec(renderer_spec)
renderer_spec.loader.exec_module(renderer_module)
nested_box_renderer = renderer_module.nested_box_renderer


class NestedBoxTemplate(PDFBaseUtils):
    """Nested Box Template Handler Class"""
    
    def __init__(self, max_pages_per_file: int = 100):
        """Initialize Nested Box Template"""
        super().__init__(max_pages_per_file)
    
    def create_multi_level_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """
        Create multi-level PDF labels for nested box template

        Args:
            data: Excel数据
            params: 用户参数 (张/盒, 盒/小箱, 小箱/大箱, 选择外观)
            output_dir: 输出目录
            excel_file_path: Excel文件路径

        Returns:
            生成的文件路径字典
        """
        # 计算数量 - 三级结构：张→盒→小箱→大箱
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])  # 这个参数用于确定结束号
        small_boxes_per_large_box = int(params["小箱/大箱"])

        # 计算各级数量
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)
        total_large_boxes = math.ceil(total_small_boxes / small_boxes_per_large_box)

        # 创建输出目录
        clean_theme = data['主题'].replace('\n', ' ').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('*', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').replace('!', '_')
        folder_name = f"{data['客户编码']}+{clean_theme}+标签"
        full_output_dir = Path(output_dir) / folder_name
        full_output_dir.mkdir(parents=True, exist_ok=True)

        generated_files = {}

        # 生成套盒模板的盒标 - 第二个参数用于结束号
        selected_appearance = params["选择外观"]
        box_label_path = full_output_dir / f"{data['客户编码']}+{clean_theme}+套盒盒标.pdf"

        self._create_nested_box_label(data, params, str(box_label_path), selected_appearance, excel_file_path)
        generated_files["盒标"] = str(box_label_path)

        # 生成套盒模板小箱标
        small_box_path = full_output_dir / f"{data['客户编码']}+{clean_theme}+套盒小箱标.pdf"
        self._create_nested_small_box_label(
            data, params, str(small_box_path), excel_file_path
        )
        generated_files["小箱标"] = str(small_box_path)

        # 生成套盒模板大箱标
        large_box_path = full_output_dir / f"{data['客户编码']}+{clean_theme}+套盒大箱标.pdf"
        self._create_nested_large_box_label(
            data, params, str(large_box_path), excel_file_path
        )
        generated_files["大箱标"] = str(large_box_path)

        return generated_files

    def _create_nested_box_label(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, style: str, excel_file_path: str = None
    ):
        """创建套盒模板的盒标 - 基于Excel文件的开始号和结束号"""
        # 分析Excel文件获取套盒特有的数据
        excel_path = excel_file_path
        print(f"🔍 正在分析套盒模板Excel文件: {excel_path}")
        
        # 使用套盒模板专属数据处理器
        excel_data = nested_box_data_processor.extract_box_label_data(excel_path)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        
        # 套盒模板参数分析
        pieces_per_box = int(params["张/盒"])
        boxes_per_ending_unit = int(params["盒/小箱"])  # 在套盒模板中，这个参数用于结束号的范围计算
        group_size = int(params["小箱/大箱"])
        
        print(f"✅ 套盒模板参数:")
        print(f"   张/盒: {pieces_per_box}")
        print(f"   盒/小箱(结束号范围): {boxes_per_ending_unit}")
        print(f"   小箱/大箱(分组大小): {group_size}")
        
        # 解析开始号的格式
        import re
        start_match = re.search(r'(.+?)(\d+)-(\d+)', base_number)
        
        if start_match:
            start_prefix = start_match.group(1)
            start_main = int(start_match.group(2))
            start_suffix = int(start_match.group(3))
            
            print(f"✅ 解析序列号格式:")
            print(f"   开始: {start_prefix}{start_main:05d}-{start_suffix:02d}")
            print(f"   结束范围由用户参数控制: {boxes_per_ending_unit}")
            
        else:
            print("⚠️ 无法解析序列号格式，使用默认逻辑")
            start_prefix = "JAW"
            start_main = 1001
            start_suffix = 1
        
        # 计算需要生成的盒标数量
        total_pieces = int(float(data["总张数"]))
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 创建PDF
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        c.setTitle("套盒模板盒标")
        
        width, height = self.page_size
        blank_height = height / 5
        top_text_y = height - 1.5 * blank_height
        serial_number_y = height - 3.5 * blank_height
        
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)
        
        # 生成套盒盒标 - 基于开始号到结束号的范围
        print(f"📝 开始生成套盒盒标，预计生成 {total_boxes} 个标签")
        
        for box_num in range(1, total_boxes + 1):
            if box_num > 1:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 套盒模板序列号生成逻辑 - 基于开始号和结束号范围
            box_index = box_num - 1
            
            # 计算当前盒的序列号在范围内的位置
            main_offset = box_index // boxes_per_ending_unit
            suffix_in_range = (box_index % boxes_per_ending_unit) + start_suffix
            
            current_main = start_main + main_offset
            current_number = f"{start_prefix}{current_main:05d}-{suffix_in_range:02d}"
            
            print(f"📝 生成套盒盒标 #{box_num}: {current_number}")
            
            # 渲染外观
            if style == "外观一":
                nested_box_renderer.render_nested_appearance_one(c, width, theme_text, current_number, top_text_y, serial_number_y)
            else:
                nested_box_renderer.render_nested_appearance_two(c, width, theme_text, current_number, top_text_y, serial_number_y)

        c.save()
        print(f"✅ 套盒模板盒标PDF已生成: {output_path}")

    def _create_nested_small_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        excel_file_path: str = None,
    ):
        """创建套盒模板的小箱标 - 借鉴分盒模板的计算逻辑"""
        # 获取Excel数据 - 使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用套盒模板专属数据处理器
        excel_data = nested_box_data_processor.extract_small_box_label_data(excel_path)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        remark_text = excel_data.get('客户编码') or 'Unknown Client'
        
        # 套盒模板不需要复杂的分组逻辑，直接使用简化逻辑
        
        # 计算参数
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        pieces_per_small_box = pieces_per_box * boxes_per_small_box
        
        # 计算小箱数量
        total_pieces = int(float(data["总张数"]))
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)
        
        # 创建PDF
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"套盒小箱标-1到{total_small_boxes}")
        c.setSubject("Taohebox Small Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成指定范围的套盒小箱标
        for small_box_num in range(1, total_small_boxes + 1):
            if small_box_num > 1:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 计算套盒模板的序列号范围 - 简化逻辑
            import re
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取第一个数字（主号）的起始位置
                digit_start = match.start()
                # 截取主号前面的所有字符作为前缀
                prefix_part = base_number[:digit_start]
                base_main_num = int(match.group(1))  # 主号
                
                # 套盒模板小箱标的简化逻辑：
                # 每个小箱标对应一个主号，包含连续的boxes_per_small_box个副号
                # 第1个小箱：主号base_main_num，副号01-06
                # 第2个小箱：主号base_main_num+1，副号01-06
                # 第3个小箱：主号base_main_num+2，副号01-06
                
                current_main_number = base_main_num + (small_box_num - 1)  # 当前小箱对应的主号
                
                # 副号始终从01开始，到boxes_per_small_box结束
                start_suffix = 1
                end_suffix = boxes_per_small_box
                
                start_serial = f"{prefix_part}{current_main_number:05d}-{start_suffix:02d}"
                end_serial = f"{prefix_part}{current_main_number:05d}-{end_suffix:02d}"
                
                # 套盒小箱标显示序列号范围
                serial_range = f"{start_serial}-{end_serial}"
            else:
                serial_range = f"DSK{small_box_num:05d}-DSK{small_box_num:05d}"

            print(f"📝 小箱标 #{small_box_num}: 主号{current_main_number}, 副号{start_suffix}-{end_suffix} = {serial_range}")

            # 计算套盒小箱标的Carton No（简单的小箱编号）
            carton_no = str(small_box_num)

            # 绘制套盒小箱标表格
            nested_box_renderer.draw_nested_small_box_table(c, width, height, theme_text, pieces_per_small_box, 
                                                             serial_range, carton_no, remark_text)

        c.save()
        print(f"✅ 套盒模板小箱标PDF已生成: {output_path}")

    def _create_nested_large_box_label(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, excel_file_path: str = None):
        """创建套盒模板的大箱标"""
        # 获取Excel数据 - 使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用套盒模板专属数据处理器
        excel_data = nested_box_data_processor.extract_large_box_label_data(excel_path)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        remark_text = excel_data.get('客户编码') or 'Unknown Client'
        
        # 获取参数
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        small_boxes_per_large_box = int(params["小箱/大箱"])
        
        print(f"✅ 套盒大箱标参数: 张/盒={pieces_per_box}, 盒/小箱={boxes_per_small_box}, 小箱/大箱={small_boxes_per_large_box}")
        
        # 计算每小箱和每大箱的数量
        pieces_per_small_box = pieces_per_box * boxes_per_small_box
        pieces_per_large_box = pieces_per_small_box * small_boxes_per_large_box
        
        print(f"✅ 计算结果: 每小箱{pieces_per_small_box}PCS, 每大箱{pieces_per_large_box}PCS")
        
        # 计算大箱数量
        total_pieces = int(float(data["总张数"]))
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)
        total_large_boxes = math.ceil(total_small_boxes / small_boxes_per_large_box)
        
        # 创建PDF
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF属性
        c.setPageCompression(1)
        c.setTitle(f"套盒大箱标-1到{total_large_boxes}")
        c.setSubject("Taohebox Large Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成大箱标
        for large_box_num in range(1, total_large_boxes + 1):
            if large_box_num > 1:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 计算当前大箱包含的小箱范围
            start_small_box = (large_box_num - 1) * small_boxes_per_large_box + 1
            end_small_box = start_small_box + small_boxes_per_large_box - 1
            
            # 计算序列号范围 - 从第一个小箱的起始号到最后一个小箱的结束号
            import re
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取第一个数字（主号）的起始位置
                digit_start = match.start()
                # 截取主号前面的所有字符作为前缀
                prefix_part = base_number[:digit_start]
                base_main_num = int(match.group(1))  # 主号
                
                # 第一个小箱的序列号范围
                first_main_number = base_main_num + (start_small_box - 1)
                first_start_serial = f"{prefix_part}{first_main_number:05d}-01"
                
                # 最后一个小箱的序列号范围
                last_main_number = base_main_num + (end_small_box - 1)
                last_end_serial = f"{prefix_part}{last_main_number:05d}-{boxes_per_small_box:02d}"
                
                # 大箱标显示完整序列号范围
                serial_range = f"{first_start_serial}-{last_end_serial}"
            else:
                serial_range = f"DSK{large_box_num:05d}-DSK{large_box_num:05d}"

            print(f"📝 大箱标 #{large_box_num}: 包含小箱{start_small_box}-{end_small_box}, 序列号范围={serial_range}")

            # 计算套盒大箱标的Carton No（小箱范围格式）
            carton_range = f"{start_small_box}-{end_small_box}"

            # 绘制套盒大箱标表格
            nested_box_renderer.draw_nested_large_box_table(c, width, height, theme_text, pieces_per_large_box, 
                                                             serial_range, carton_range, remark_text)

        c.save()
        print(f"✅ 套盒模板大箱标PDF已生成: {output_path}")




