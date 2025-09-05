"""
Split Box Template - Multi-level PDF generation with special serial number logic
"""
import math
from pathlib import Path
from typing import Dict, Any
from reportlab.pdfgen import canvas
from reportlab.lib.colors import CMYKColor
# 导入基础工具类
from src.utils.pdf_base import PDFBaseUtils

# 导入分盒模板专属数据处理器和渲染器
from src.pdf.split_box.data_processor import split_box_data_processor
from src.pdf.split_box.renderer import split_box_renderer


class SplitBoxTemplate(PDFBaseUtils):
    """Split Box Template Handler Class"""
    
    def __init__(self, max_pages_per_file: int = 100):
        """Initialize Split Box Template"""
        super().__init__(max_pages_per_file)
    
    def create_multi_level_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """
        Create multi-level PDF labels for split box template

        Args:
            data: Excel数据
            params: 用户参数 (张/盒, 盒/小箱, 小箱/大箱, 选择外观)
            output_dir: 输出目录

        Returns:
            生成的文件路径字典
        """
        # 计算数量 - 三级结构：张→盒→小箱→大箱
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        small_boxes_per_large_box = int(params["小箱/大箱"])

        # 计算各级数量 (使用向上取整处理余数)
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)
        total_large_boxes = math.ceil(total_small_boxes / small_boxes_per_large_box)

        # 创建输出目录
        clean_theme = data['主题'].replace('\n', ' ').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('*', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').replace('!', '_')
        folder_name = f"{data['客户编码']}+{clean_theme}+标签"
        full_output_dir = Path(output_dir) / folder_name
        full_output_dir.mkdir(parents=True, exist_ok=True)

        generated_files = {}

        # 生成盒标 (只生成用户选择的外观)
        selected_appearance = params["选择外观"]
        box_label_path = (
            full_output_dir
            / f"{data['客户编码']}+{clean_theme}+分盒盒标+{selected_appearance}.pdf"
        )

        self._create_split_box_label(data, params, str(box_label_path), selected_appearance, excel_file_path)
        generated_files["盒标"] = str(box_label_path)

        # 生成小箱标
        small_box_path = (
            full_output_dir / f"{data['客户编码']}+{clean_theme}+分盒小箱标.pdf"
        )
        remainder_info = {"total_boxes": total_boxes}
        self._create_split_box_small_box_label(
            data, params, str(small_box_path), total_small_boxes, remainder_info, excel_file_path
        )
        generated_files["小箱标"] = str(small_box_path)

        # 生成大箱标
        large_box_path = (
            full_output_dir / f"{data['客户编码']}+{clean_theme}+分盒大箱标.pdf"
        )
        self._create_split_box_large_box_label(
            data, params, str(large_box_path), total_large_boxes, excel_file_path
        )
        generated_files["大箱标"] = str(large_box_path)

        return generated_files

    def _create_split_box_label(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, style: str, excel_file_path: str = None):
        """创建split box template box labels - 特殊序列号逻辑"""
        # 计算总盒数
        total_pieces = int(float(data["总张数"]))  # 处理Excel的float值
        pieces_per_box = int(params["张/盒"])
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 获取盒标内容 - 只使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        # 使用分盒模板专属数据处理器
        excel_data = split_box_data_processor.extract_box_label_data(excel_path)
        top_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        
        # 从用户输入的第三个参数获取分组大小（从"小箱/大箱"参数获取）
        try:
            group_size = int(params["小箱/大箱"])  # 用户的第三个参数，控制副号满几进一
            if group_size <= 0:  # 避免除零错误
                group_size = 2
            print(f"✅ 分盒盒标使用用户输入分组大小: {group_size} (小箱/大箱)")
        except (ValueError, KeyError) as e:
            print(f"⚠️ 获取小箱/大箱参数失败: {e}")
            group_size = 2  # 默认分组大小
        
        # 直接创建单个PDF文件，包含所有盒标（移除分页限制）
        self._create_single_split_box_label_file(
            data, params, output_path, style, 
            1, total_boxes, top_text, base_number, group_size
        )

    def _create_single_split_box_label_file(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, 
                                           style: str, start_box: int, end_box: int, top_text: str, base_number: str, group_size: int):
        """创建单个分盒模板盒标PDF文件"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"分盒盒标-{style}-{start_box}到{end_box}")
        c.setSubject("Fenhe Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)
        
        # 真正的三等分留白布局：每个留白区域高度相等
        blank_height = height / 5  # 每个留白区域高度：10mm
        
        # 布局位置计算（确保三个留白区域等高）
        top_text_y = height - 1.5 * blank_height      # 产品名称居中在区域2
        serial_number_y = height - 3.5 * blank_height # 序列号居中在区域4

        # 生成指定范围的盒标
        for box_num in range(start_box, end_box + 1):
            if box_num > start_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 生成分盒模板的序列号 - 特殊处理主号和副号
            import re
            # 分盒模板格式：前缀+主号+"-"+副号，如 ABC01001-02
            # 需要找到主号（第一个数字）前面的字符作为前缀
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取第一个数字（主号）的起始位置
                main_digit_start = match.start()
                # 截取主号前面的所有字符作为前缀
                prefix_part = base_number[:main_digit_start]  # 比如 "ABC", "GLA-", "A—B@C"
                base_main_num = int(match.group(1))  # 主号，如 01001
                
                # 计算主号码和后缀
                # box_num从1开始，需要转换为0基数来计算
                box_index = box_num - 1
                main_number_offset = box_index // group_size  # 主号码偏移
                suffix_number = (box_index % group_size) + 1  # 后缀号码(1开始)
                
                new_main_number = base_main_num + main_number_offset
                # 分盒模板格式：前缀 + 主号码 + "-" + 副号
                current_number = f"{prefix_part}{new_main_number:05d}-{suffix_number:02d}"
            else:
                # 备用方案
                current_number = f"DSK{box_num:05d}-01"
            
            # 分盒模板只有一种固定外观，使用简洁标准样式
            split_box_renderer.render_appearance_one(c, width, top_text, current_number, top_text_y, serial_number_y)

        c.save()


    def _create_split_box_small_box_label(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, 
                                     total_small_boxes: int, remainder_info: Dict[str, Any], excel_file_path: str = None):
        """创建split box template small box labels"""
        # 获取Excel数据 - 使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用分盒模板专属数据处理器
        excel_data = split_box_data_processor.extract_small_box_label_data(excel_path)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        remark_text = excel_data.get('客户编码') or 'Unknown Client'
        
        # 获取用户输入的分组大小（从"小箱/大箱"参数获取）
        try:
            group_size = int(params["小箱/大箱"])  # 用户的第三个参数，控制副号满几进一
            if group_size <= 0:
                group_size = 2
            print(f"✅ 分盒小箱标使用用户输入分组大小: {group_size} (小箱/大箱)")
        except (ValueError, KeyError) as e:
            print(f"⚠️ 获取小箱/大箱参数失败: {e}")
            group_size = 2  # 默认分组大小
        
        # 计算参数
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        pieces_per_small_box = pieces_per_box * boxes_per_small_box
        
        # 直接创建单个PDF文件，包含所有小箱标
        self._create_single_split_box_small_box_label_file(
            data, params, output_path, 1, total_small_boxes,
            theme_text, base_number, remark_text, pieces_per_small_box, 
            boxes_per_small_box, total_small_boxes, group_size
        )

    def _create_single_split_box_small_box_label_file(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
                                                 start_small_box: int, end_small_box: int, theme_text: str, base_number: str,
                                                 remark_text: str, pieces_per_small_box: int, boxes_per_small_box: int, 
                                                 total_small_boxes: int, group_size: int):
        """创建单个分盒小箱标PDF文件"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"分盒小箱标-{start_small_box}到{end_small_box}")
        c.setSubject("Fenhe Small Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成指定范围的分盒小箱标
        for small_box_num in range(start_small_box, end_small_box + 1):
            if small_box_num > start_small_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 计算分盒模板的序列号范围
            import re
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取第一个数字（主号）的起始位置
                digit_start = match.start()
                # 截取主号前面的所有字符作为前缀
                prefix_part = base_number[:digit_start]
                base_main_num = int(match.group(1))  # 主号
                
                # 分盒模板小箱标的特殊逻辑：
                # 每个小箱标包含boxes_per_small_box个盒标的序列号范围
                # 计算当前小箱标包含的盒标范围
                start_box_index = (small_box_num - 1) * boxes_per_small_box  # 起始盒标索引(0基数)
                end_box_index = start_box_index + boxes_per_small_box - 1    # 结束盒标索引(0基数)
                
                # 计算起始盒标的序列号
                start_main_offset = start_box_index // group_size
                start_suffix = (start_box_index % group_size) + 1
                start_main_number = base_main_num + start_main_offset
                start_serial = f"{prefix_part}{start_main_number:05d}-{start_suffix:02d}"
                
                # 计算结束盒标的序列号
                end_main_offset = end_box_index // group_size
                end_suffix = (end_box_index % group_size) + 1
                end_main_number = base_main_num + end_main_offset
                end_serial = f"{prefix_part}{end_main_number:05d}-{end_suffix:02d}"
                
                # 分盒小箱标显示序列号范围
                serial_range = f"{start_serial}-{end_serial}"
            else:
                serial_range = f"DSK{small_box_num:05d}-DSK{small_box_num:05d}"

            # 计算分盒小箱标的Carton No（主箱号-副箱号格式）
            main_box_num = ((small_box_num - 1) // group_size) + 1  # 主箱号
            sub_box_num = ((small_box_num - 1) % group_size) + 1    # 副箱号
            carton_no = f"{main_box_num}-{sub_box_num}"

            # 绘制分盒小箱标表格
            split_box_renderer.draw_split_box_small_box_table(c, width, height, theme_text, pieces_per_small_box, 
                                           serial_range, carton_no, remark_text)

        c.save()


    def _create_split_box_large_box_label(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, 
                                     total_large_boxes: int, excel_file_path: str = None):
        """创建split box template large box labels - 完全参考小箱标模式"""
        # 获取Excel数据 - 使用关键字提取，与小箱标相同
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用分盒模板专属数据处理器
        excel_data = split_box_data_processor.extract_large_box_label_data(excel_path)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        remark_text = excel_data.get('客户编码') or 'Unknown Client'
        
        # 获取用户输入的分组大小（从"小箱/大箱"参数获取）
        try:
            group_size = int(params["小箱/大箱"])  # 用户的第三个参数，控制副号满几进一
            print(f"✅ 分盒大箱标使用用户输入分组大小: {group_size} (小箱/大箱)")
        except (ValueError, KeyError) as e:
            print(f"⚠️ 获取小箱/大箱参数失败: {e}")
            group_size = 2  # 默认分组大小
        
        # 计算参数 - 大箱标专用
        pieces_per_box = int(params["张/盒"])  # 第一个参数：张/盒
        boxes_per_small_box = int(params["盒/小箱"])  # 第二个参数：盒/小箱
        small_boxes_per_large_box = int(params["小箱/大箱"])  # 第三个参数：小箱/大箱
        
        # 直接创建单个PDF文件，包含所有大箱标
        self._create_single_split_box_large_box_label_file(
            data, params, output_path, 1, total_large_boxes,
            theme_text, base_number, remark_text, pieces_per_box, 
            small_boxes_per_large_box, total_large_boxes, group_size
        )

    def _create_single_split_box_large_box_label_file(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
                                                 start_large_box: int, end_large_box: int, theme_text: str, base_number: str,
                                                 remark_text: str, pieces_per_box: int, small_boxes_per_large_box: int, 
                                                 total_large_boxes: int, group_size: int):
        """创建单个分盒大箱标PDF文件 - 完全参考小箱标"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"分盒大箱标-{start_large_box}到{end_large_box}")
        c.setSubject("Fenhe Large Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成指定范围的大箱标
        for large_box_num in range(start_large_box, end_large_box + 1):
            if large_box_num > start_large_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 计算当前大箱的序列号范围
            import re
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取第一个数字（主号）的起始位置
                main_digit_start = match.start()
                # 截取主号前面的所有字符作为前缀
                prefix_part = base_number[:main_digit_start]  # 比如 "LGM"
                base_main_num = int(match.group(1))  # 主号，如 01001
                
                # 计算当前大箱的主号 (每个大箱主号递增)
                current_main_number = base_main_num + (large_box_num - 1)
                
                # 生成序列号范围（从01到group_size）
                start_serial = f"{prefix_part}{current_main_number:05d}-01"
                end_serial = f"{prefix_part}{current_main_number:05d}-{group_size:02d}"
                serial_range = f"{start_serial}-{end_serial}"
            else:
                # 备用方案
                serial_range = f"DSK{large_box_num:05d}-01-DSK{large_box_num:05d}-{group_size:02d}"
            
            # 绘制大箱标表格 - 完全使用小箱标的表格结构
            split_box_renderer.draw_split_box_large_box_table(c, width, height, theme_text, pieces_per_box, 
                                           small_boxes_per_large_box, serial_range, 
                                           str(large_box_num), remark_text)

        c.save()

