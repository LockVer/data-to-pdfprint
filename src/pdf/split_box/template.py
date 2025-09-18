"""
Split Box Template - Multi-level PDF generation with special serial number logic
"""
import math
from datetime import datetime
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
    
    def __init__(self):
        """Initialize Split Box Template"""
        super().__init__()
    
    def create_multi_level_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """
        Create multi-level PDF labels for split box template

        Args:
            data: Excel数据
            params: 用户参数 (张/盒, 盒/套, 盒/小箱, 小箱/大箱, 选择外观, 是否有小箱)
            output_dir: 输出目录

        Returns:
            生成的文件路径字典
        """
        # 🔧 临时修复：确保"盒/套"参数存在
        print(f"🔍 主入口调试：params内容 = {params}")
        if "盒/套" not in params:
            print("⚠️ 警告：缺少'盒/套'参数，使用默认值1")
            params["盒/套"] = 1
        # 检查是否有小箱
        has_small_box = params.get("是否有小箱", True)
        
        if has_small_box:
            return self._create_three_level_pdfs(data, params, output_dir, excel_file_path)
        else:
            return self._create_two_level_pdfs(data, params, output_dir, excel_file_path)
    
    def _create_three_level_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """
        创建三级包装的PDF（有小箱）
        """
        # 计算数量 - 三级结构：张→盒→小箱→大箱
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        small_boxes_per_large_box = int(params["小箱/大箱"])

        # 🔧 修复：计算各级数量，基于套数计算而非简单除法
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 基于套数的正确计算
        boxes_per_set = int(params.get("盒/套", params.get("boxes_per_set", 1)))
        total_sets = math.ceil(total_boxes / boxes_per_set)
        
        # 计算每套需要的小箱数和大箱数
        small_boxes_per_set = math.ceil(boxes_per_set / boxes_per_small_box)
        large_boxes_per_set = math.ceil(small_boxes_per_set / small_boxes_per_large_box)
        
        # 基于套数计算总数量
        total_small_boxes = total_sets * small_boxes_per_set
        total_large_boxes = total_sets * large_boxes_per_set
        
        print(f"🔍 三级模式数量计算 (修复后的正确逻辑):")
        print(f"    总张数: {total_pieces}, 张/盒: {pieces_per_box}")
        print(f"    总盒数: {total_boxes} = ceil({total_pieces} ÷ {pieces_per_box})")
        print(f"    盒/套: {boxes_per_set}, 总套数: {total_sets} = ceil({total_boxes} ÷ {boxes_per_set})")
        print(f"    每套小箱数: {small_boxes_per_set} = ceil({boxes_per_set} ÷ {boxes_per_small_box})")
        print(f"    每套大箱数: {large_boxes_per_set} = ceil({small_boxes_per_set} ÷ {small_boxes_per_large_box})")
        print(f"    总小箱数: {total_small_boxes} = {total_sets} × {small_boxes_per_set}")
        print(f"    总大箱数: {total_large_boxes} = {total_sets} × {large_boxes_per_set}")

        # 创建输出目录
        clean_theme = data['标签名称'].replace('\n', ' ').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('*', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').replace('!', '_')
        folder_name = f"{data['客户名称编码']}+{clean_theme}+标签"
        full_output_dir = Path(output_dir) / folder_name
        full_output_dir.mkdir(parents=True, exist_ok=True)

        # 获取参数和日期时间戳
        chinese_name = params.get("中文名称", "")
        english_name = clean_theme  # 英文名称使用清理后的主题
        customer_code = data['客户名称编码']  # 客户编号
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        generated_files = {}

        # 检查是否需要生成盒标
        has_box_label = params.get("是否有盒标", False)
        
        if has_box_label:
            # 生成分盒盒标 (分盒模板无外观选择)
            selected_appearance = params["选择外观"]  # 保留参数传递，但文件名不使用
            # 文件名格式：客户编号_中文名称_英文名称_分盒盒标_日期时间戳
            box_label_filename = f"{customer_code}_{chinese_name}_{english_name}_分盒盒标_{timestamp}.pdf"
            box_label_path = full_output_dir / box_label_filename

            self._create_split_box_label(data, params, str(box_label_path), selected_appearance, excel_file_path)
            generated_files["盒标"] = str(box_label_path)
        else:
            print("⏭️ 用户选择无盒标，跳过盒标生成")

        # 生成小箱标
        # 文件名格式：客户编号_中文名称_英文名称_分盒小箱标_日期时间戳
        small_box_filename = f"{customer_code}_{chinese_name}_{english_name}_分盒小箱标_{timestamp}.pdf"
        small_box_path = full_output_dir / small_box_filename
        remainder_info = {"total_boxes": total_boxes}
        self._create_split_box_small_box_label(
            data, params, str(small_box_path), total_small_boxes, remainder_info, excel_file_path
        )
        generated_files["小箱标"] = str(small_box_path)

        # 生成大箱标
        # 文件名格式：客户编号_中文名称_英文名称_分盒大箱标_日期时间戳
        large_box_filename = f"{customer_code}_{chinese_name}_{english_name}_分盒大箱标_{timestamp}.pdf"
        large_box_path = full_output_dir / large_box_filename
        self._create_split_box_large_box_label(
            data, params, str(large_box_path), total_large_boxes, excel_file_path
        )
        generated_files["大箱标"] = str(large_box_path)

        return generated_files
    
    def _create_two_level_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """
        创建二级包装的PDF（无小箱）
        """
        # 计算数量 - 二级结构：张→盒→箱
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        # 🔧 修复：二级模式下的参数映射问题
        # 在UI的"无小箱"模式下，"盒/小箱"实际存储的是"盒/大箱"的值
        has_small_box = params.get("是否有小箱", True)
        if has_small_box:
            # 三级模式：盒/箱 = 盒/小箱 × 小箱/大箱
            boxes_per_small_box = int(params["盒/小箱"]) 
            small_boxes_per_large_box = int(params["小箱/大箱"])
            boxes_per_large_box = boxes_per_small_box * small_boxes_per_large_box
            print(f"✅ 三级模式计算: 盒/大箱 = 盒/小箱({boxes_per_small_box}) × 小箱/大箱({small_boxes_per_large_box}) = {boxes_per_large_box}")
        else:
            # 二级模式：直接使用"盒/小箱"中存储的"盒/大箱"值
            boxes_per_large_box = int(params["盒/小箱"])  # 注意：这里存储的实际是"盒/大箱"
            print(f"✅ 二级模式计算: 盒/大箱 = {boxes_per_large_box} (直接从UI获取)")

        # 计算各级数量
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 🔧 修复：大箱数应该基于套数计算，而不是简单的总盒数除法
        boxes_per_set = int(params.get("盒/套", params.get("boxes_per_set", 1)))
        total_sets = math.ceil(total_boxes / boxes_per_set)
        large_boxes_per_set = math.ceil(boxes_per_set / boxes_per_large_box)  # 每套需要多少大箱
        total_large_boxes = total_sets * large_boxes_per_set  # 总套数 × 每套大箱数
        
        print(f"🔍 二级模式数量计算 (修复后的正确逻辑):")
        print(f"    总张数: {total_pieces}")
        print(f"    张/盒: {pieces_per_box}")
        print(f"    总盒数: {total_boxes} = ceil({total_pieces} ÷ {pieces_per_box}) = ceil({total_pieces / pieces_per_box})")
        print(f"    盒/套: {boxes_per_set}")
        print(f"    总套数: {total_sets} = ceil({total_boxes} ÷ {boxes_per_set}) = ceil({total_boxes / boxes_per_set})")
        print(f"    盒/大箱: {boxes_per_large_box}")
        print(f"    每套大箱数: {large_boxes_per_set} = ceil({boxes_per_set} ÷ {boxes_per_large_box}) = ceil({boxes_per_set / boxes_per_large_box})")
        print(f"    总大箱数: {total_large_boxes} = {total_sets} × {large_boxes_per_set} (套数 × 每套大箱数)")
        print(f"✅ 修复：使用基于套数的计算方式，而非简单的总盒数除法")

        # 创建输出目录
        clean_theme = data['标签名称'].replace('\n', ' ').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('*', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').replace('!', '_')
        folder_name = f"{data['客户名称编码']}+{clean_theme}+标签"
        full_output_dir = Path(output_dir) / folder_name
        full_output_dir.mkdir(parents=True, exist_ok=True)

        # 获取参数和日期时间戳
        chinese_name = params.get("中文名称", "")
        english_name = clean_theme  # 英文名称使用清理后的主题
        customer_code = data['客户名称编码']  # 客户编号
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        generated_files = {}

        # 检查是否需要生成盒标
        has_box_label = params.get("是否有盒标", False)
        
        if has_box_label:
            # 生成分盒盒标 (分盒模板无外观选择)
            selected_appearance = params["选择外观"]  # 保留参数传递，但文件名不使用
            # 文件名格式：客户编号_中文名称_英文名称_分盒盒标_日期时间戳
            box_label_filename = f"{customer_code}_{chinese_name}_{english_name}_分盒盒标_{timestamp}.pdf"
            box_label_path = full_output_dir / box_label_filename

            self._create_split_box_label(data, params, str(box_label_path), selected_appearance, excel_file_path)
            generated_files["盒标"] = str(box_label_path)
        else:
            print("⏭️ 用户选择无盒标，跳过盒标生成")

        # 生成箱标（复用大箱标逻辑但文件名为箱标）
        # 文件名格式：客户编号_中文名称_英文名称_分盒箱标_日期时间戳
        large_box_filename = f"{customer_code}_{chinese_name}_{english_name}_分盒箱标_{timestamp}.pdf"
        large_box_path = full_output_dir / large_box_filename
        
        self._create_two_level_large_box_label(
            data, params, str(large_box_path), total_large_boxes, total_boxes, boxes_per_large_box, excel_file_path
        )
        generated_files["箱标"] = str(large_box_path)

        return generated_files

    def _create_split_box_label(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, style: str, excel_file_path: str = None):
        """创建split box template box labels - 特殊序列号逻辑"""
        # 计算总盒数
        total_pieces = int(float(data["总张数"]))  # 处理Excel的float值
        pieces_per_box = int(params["张/盒"])
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 使用统一数据处理后的标准四字段（优先使用传入的data参数）
        top_text = data.get('标签名称') or 'Unknown Title'
        base_number = data.get('开始号') or 'DEFAULT01001'
        print(f"✅ 分盒盒标使用统一数据: 主题='{top_text}', 开始号='{base_number}'")
        
        # 获取用户输入的包装参数
        boxes_per_small_box = int(params["盒/小箱"])
        small_boxes_per_large_box = int(params["小箱/大箱"])
        print(f"✅ 分盒盒标参数: 盒/小箱={boxes_per_small_box}, 小箱/大箱={small_boxes_per_large_box}")
        
        # 直接创建单个PDF文件，包含所有盒标（移除分页限制）
        self._create_single_split_box_label_file(
            data, params, output_path, style, 
            1, total_boxes, top_text, base_number, boxes_per_small_box, small_boxes_per_large_box
        )

    def _create_single_split_box_label_file(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, 
                                           style: str, start_box: int, end_box: int, top_text: str, base_number: str, boxes_per_small_box: int, small_boxes_per_large_box: int):
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

        # 获取中文名称用于空白首页
        chinese_name = params.get("中文名称", "")
        
        # 生成指定范围的盒标
        for box_num in range(start_box, end_box + 1):
            # 🔥 新增：在第一个标签时添加空白首页（只对外观1）
            if box_num == start_box and style == "外观一" and chinese_name:
                print(f"📝 生成分盒盒标空白首页: {chinese_name}")
                split_box_renderer.render_blank_first_page(c, width, height, chinese_name)
                c.showPage()
                c.setFillColor(cmyk_black)
            
            if box_num > start_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 使用数据处理器生成序列号
            current_number = split_box_data_processor.generate_split_box_serial_number(
                base_number, box_num, boxes_per_small_box, small_boxes_per_large_box
            )
            
            # 分盒模板只有一种固定外观，使用简洁标准样式
            split_box_renderer.render_appearance_one(c, width, top_text, current_number, top_text_y, serial_number_y)

        c.save()


    def _create_split_box_small_box_label(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, 
                                     total_small_boxes: int, remainder_info: Dict[str, Any], excel_file_path: str = None):
        """创建split box template small box labels"""
        # 获取Excel数据 - 使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用统一数据处理后的标准四字段（优先使用传入的data参数）
        theme_text = data.get('标签名称') or 'Unknown Title'
        base_number = data.get('开始号') or 'DEFAULT01001'
        remark_text = data.get('客户名称编码') or 'Unknown Client'
        print(f"✅ 分盒小箱标使用统一数据: 主题='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}'")
        
        # 获取用户输入的包装参数
        print(f"🔍 调试：params内容 = {params}")
        print(f"🔍 调试：params.keys() = {list(params.keys())}")
        print(f"🔍 调试：'盒/套' in params = {'盒/套' in params}")
        if '盒/套' in params:
            print(f"🔍 调试：params['盒/套'] = {params['盒/套']}")
        
        boxes_per_set = int(params.get("盒/套", params.get("boxes_per_set", 1)))  # 兼容处理
        print(f"🔍 调试：最终 boxes_per_set = {boxes_per_set}")
        boxes_per_small_box = int(params["盒/小箱"])
        small_boxes_per_large_box = int(params["小箱/大箱"])
        serial_font_size = int(params.get("序列号字体大小", 10))
        print(f"✅ 分盒小箱标参数: 盒/套={boxes_per_set}, 盒/小箱={boxes_per_small_box}, 小箱/大箱={small_boxes_per_large_box}, 序列号字体大小={serial_font_size}")
        
        # 计算参数
        pieces_per_box = int(params["张/盒"])
        pieces_per_small_box = pieces_per_box * boxes_per_small_box
        
        # 从remainder_info获取total_boxes
        total_boxes = remainder_info.get("total_boxes", 0)
        
        # 直接创建单个PDF文件，包含所有小箱标
        self._create_single_split_box_small_box_label_file(
            data, params, output_path, 1, total_small_boxes,
            theme_text, base_number, remark_text, pieces_per_small_box, 
            boxes_per_set, boxes_per_small_box, total_small_boxes, small_boxes_per_large_box, total_boxes, serial_font_size
        )

    def _create_single_split_box_small_box_label_file(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
                                                 start_small_box: int, end_small_box: int, theme_text: str, base_number: str,
                                                 remark_text: str, pieces_per_small_box: int, boxes_per_set: int, boxes_per_small_box: int, 
                                                 total_small_boxes: int, small_boxes_per_large_box: int, total_boxes: int, serial_font_size: int = 10):
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

        # 在第一页添加空箱标签（仅在处理第一个小箱时）
        if start_small_box == 1:
            # 获取中文名称参数
            chinese_name = params.get("中文名称", "")
            # 获取标签模版类型
            template_type = params.get("标签模版", "有纸卡备注")
            
            # 根据标签模版类型选择空箱标签渲染函数
            if template_type == "有纸卡备注":
                split_box_renderer.render_empty_box_label(c, width, height, chinese_name, remark_text)
            else:  # "无纸卡备注"
                split_box_renderer.render_empty_box_label_no_paper_card(c, width, height, chinese_name, remark_text)
            
            c.showPage()
            c.setFillColor(cmyk_black)

        # 生成指定范围的分盒小箱标
        for small_box_num in range(start_small_box, end_small_box + 1):
            if small_box_num > start_small_box or start_small_box == 1:  # 修改条件，考虑空标签页
                if not (small_box_num == start_small_box and start_small_box == 1):  # 避免重复showPage
                    c.showPage()
                    c.setFillColor(cmyk_black)

            # 🔧 使用修复后的数据处理器计算序列号范围（包含边界检查）
            serial_range = split_box_data_processor.generate_split_small_box_serial_range(
                base_number, small_box_num, boxes_per_small_box, small_boxes_per_large_box, total_boxes
            )

            # 🔧 计算当前小箱的实际张数（考虑最后一小箱的边界情况）
            pieces_per_box = int(params["张/盒"])
            # 计算当前小箱实际包含的盒数
            start_box = (small_box_num - 1) * boxes_per_small_box + 1
            end_box = min(start_box + boxes_per_small_box - 1, total_boxes)
            actual_boxes_in_small_box = end_box - start_box + 1
            actual_pieces_in_small_box = actual_boxes_in_small_box * pieces_per_box

            # 计算分盒小箱标的Carton No - 基于最新逻辑整理
            print(f"\n📦 准备计算小箱标 #{small_box_num} 的Carton No")
            carton_no = split_box_data_processor.calculate_carton_number_for_small_box(
                small_box_num, boxes_per_set, boxes_per_small_box
            )
            print(f"📦 小箱标 #{small_box_num} Carton No计算完成: {carton_no}\n")

            # 获取标签模版类型 - 参照常规模版的实现方式
            template_type = params.get("标签模版", "有纸卡备注")
            
            # 绘制分盒小箱标表格（使用实际张数，根据模版类型选择函数）
            if template_type == "有纸卡备注":
                split_box_renderer.draw_split_box_small_box_table(c, width, height, theme_text, actual_pieces_in_small_box, 
                                               serial_range, carton_no, remark_text, True, serial_font_size)
            else:  # "无纸卡备注"
                split_box_renderer.draw_split_box_small_box_table_no_paper_card(c, width, height, theme_text, actual_pieces_in_small_box, 
                                               serial_range, carton_no, remark_text, serial_font_size)

        c.save()


    def _create_split_box_large_box_label(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, 
                                     total_large_boxes: int, excel_file_path: str = None):
        """创建split box template large box labels - 完全参考小箱标模式"""
        # 获取Excel数据 - 使用关键字提取，与小箱标相同
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用统一数据处理后的标准四字段（优先使用传入的data参数）
        theme_text = data.get('标签名称') or 'Unknown Title'
        base_number = data.get('开始号') or 'DEFAULT01001'
        remark_text = data.get('客户名称编码') or 'Unknown Client'
        print(f"✅ 分盒大箱标使用统一数据: 主题='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}'")
        
        # 获取用户输入的包装参数
        print(f"🔍 调试：大箱标params内容 = {params}")
        boxes_per_set = int(params.get("盒/套", params.get("boxes_per_set", 1)))  # 兼容处理
        boxes_per_small_box = int(params["盒/小箱"])
        small_boxes_per_large_box = int(params["小箱/大箱"])
        serial_font_size = int(params.get("序列号字体大小", 10))
        print(f"✅ 分盒大箱标参数: 盒/套={boxes_per_set}, 盒/小箱={boxes_per_small_box}, 小箱/大箱={small_boxes_per_large_box}, 序列号字体大小={serial_font_size}")
        
        # 计算参数 - 大箱标专用
        pieces_per_box = int(params["张/盒"])  # 第一个参数：张/盒
        
        # 从remainder_info获取total_boxes
        total_pieces = int(float(data["总张数"]))
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 计算总套数
        total_sets = math.ceil(total_boxes / boxes_per_set)
        
        # 直接创建单个PDF文件，包含所有大箱标
        self._create_single_split_box_large_box_label_file(
            data, params, output_path, 1, total_large_boxes,
            theme_text, base_number, remark_text, pieces_per_box, 
            boxes_per_set, boxes_per_small_box, small_boxes_per_large_box, total_large_boxes, total_sets, serial_font_size
        )

    def _create_single_split_box_large_box_label_file(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
                                                 start_large_box: int, end_large_box: int, theme_text: str, base_number: str,
                                                 remark_text: str, pieces_per_box: int, boxes_per_set: int, boxes_per_small_box: int, 
                                                 small_boxes_per_large_box: int, total_large_boxes: int, total_sets: int, serial_font_size: int = 10):
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

        # 在第一页添加空箱标签（仅在处理第一个大箱时）
        if start_large_box == 1:
            # 获取中文名称参数
            chinese_name = params.get("中文名称", "")
            # 获取标签模版类型
            template_type = params.get("标签模版", "有纸卡备注")
            
            # 根据标签模版类型选择空箱标签渲染函数
            if template_type == "有纸卡备注":
                split_box_renderer.render_empty_box_label(c, width, height, chinese_name, remark_text)
            else:  # "无纸卡备注"
                split_box_renderer.render_empty_box_label_no_paper_card(c, width, height, chinese_name, remark_text)
            
            c.showPage()
            c.setFillColor(cmyk_black)

        # 生成指定范围的大箱标
        for large_box_num in range(start_large_box, end_large_box + 1):
            if large_box_num > start_large_box or start_large_box == 1:  # 修改条件，考虑空标签页
                if not (large_box_num == start_large_box and start_large_box == 1):  # 避免重复showPage
                    c.showPage()
                    c.setFillColor(cmyk_black)

            # 计算当前大箱的序列号范围，使用正确的副号进位阈值
            # 从data_processor中获取序列号范围，但需要计算总盒数边界
            total_pieces = int(float(data["总张数"]))
            total_boxes = math.ceil(total_pieces / pieces_per_box)
            
            # 使用数据处理器计算序列号范围
            serial_range = split_box_data_processor.generate_split_large_box_serial_range(
                base_number, large_box_num, small_boxes_per_large_box, boxes_per_small_box, total_boxes
            )
            
            # 计算大箱标的Carton No - 基于最新逻辑整理  
            boxes_per_large_box = boxes_per_small_box * small_boxes_per_large_box
            print(f"\n📦 准备计算大箱标 #{large_box_num} 的Carton No")
            carton_no = split_box_data_processor.calculate_carton_range_for_large_box(
                large_box_num, boxes_per_set, boxes_per_large_box, total_sets
            )
            print(f"📦 大箱标 #{large_box_num} Carton No计算完成: {carton_no}\n")
            
            # 获取标签模版类型 - 参照常规模版的实现方式
            template_type = params.get("标签模版", "有纸卡备注")
            
            # 绘制大箱标表格 - 传入完整的包装参数，根据模版类型选择函数
            if template_type == "有纸卡备注":
                split_box_renderer.draw_split_box_large_box_table(c, width, height, theme_text, pieces_per_box, 
                                               boxes_per_small_box, small_boxes_per_large_box, serial_range, 
                                               carton_no, remark_text, serial_font_size)
            else:  # "无纸卡备注"
                split_box_renderer.draw_split_box_large_box_table_no_paper_card(c, width, height, theme_text, pieces_per_box, 
                                               boxes_per_small_box, small_boxes_per_large_box, serial_range, 
                                               carton_no, remark_text, serial_font_size)

        c.save()

    def _create_two_level_large_box_label(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, 
                                     total_large_boxes: int, total_boxes: int, boxes_per_large_box: int, excel_file_path: str = None):
        """创建二级模式的箱标（无小箱）"""
        # 获取Excel数据 - 使用关键字提取，与大箱标相同
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用统一数据处理后的标准四字段（优先使用传入的data参数）
        theme_text = data.get('标签名称') or 'Unknown Title'
        base_number = data.get('开始号') or 'DEFAULT01001'
        remark_text = data.get('客户名称编码') or 'Unknown Client'
        print(f"✅ 分盒箱标使用统一数据: 主题='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}'")
        
        # 计算参数 - 箱标专用（二级模式）
        pieces_per_box = int(params["张/盒"])  # 第一个参数：张/盒
        serial_font_size = int(params.get("序列号字体大小", 10))
        print(f"✅ 分盒箱标参数: 盒/箱={boxes_per_large_box}, 序列号字体大小={serial_font_size}")
        
        # 直接创建单个PDF文件，包含所有箱标
        self._create_single_two_level_large_box_label_file(
            data, params, output_path, 1, total_large_boxes,
            theme_text, base_number, remark_text, pieces_per_box, 
            boxes_per_large_box, total_large_boxes, total_boxes, serial_font_size
        )

    def _create_single_two_level_large_box_label_file(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
                                                 start_large_box: int, end_large_box: int, theme_text: str, base_number: str,
                                                 remark_text: str, pieces_per_box: int, boxes_per_large_box: int, 
                                                 total_large_boxes: int, total_boxes: int, serial_font_size: int = 10):
        """创建单个分盒箱标PDF文件（二级模式）"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"分盒箱标-{start_large_box}到{end_large_box}")
        c.setSubject("Fenhe Box Label (Two Level)")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 在第一页添加空箱标签（仅在处理第一个箱时）
        if start_large_box == 1:
            # 获取中文名称参数
            chinese_name = params.get("中文名称", "")
            # 获取标签模版类型
            template_type = params.get("标签模版", "有纸卡备注")
            
            # 根据标签模版类型选择空箱标签渲染函数
            if template_type == "有纸卡备注":
                split_box_renderer.render_empty_box_label(c, width, height, chinese_name, remark_text)
            else:  # "无纸卡备注"
                split_box_renderer.render_empty_box_label_no_paper_card(c, width, height, chinese_name, remark_text)
            
            c.showPage()
            c.setFillColor(cmyk_black)

        # 生成指定范围的箱标
        for large_box_num in range(start_large_box, end_large_box + 1):
            if large_box_num > start_large_box or start_large_box == 1:  # 修改条件，考虑空标签页
                if not (large_box_num == start_large_box and start_large_box == 1):  # 避免重复showPage
                    c.showPage()
                    c.setFillColor(cmyk_black)

            # 计算当前箱的序列号范围，使用分盒模板的特殊逻辑
            # 二级模式：复用大箱标逻辑，但设置 small_boxes_per_large_box = 1，进位阈值 = boxes_per_large_box
            serial_range = split_box_data_processor.generate_split_large_box_serial_range(
                base_number, large_box_num, 1, boxes_per_large_box, total_boxes
            )
            
            # 🔧 计算二级模式箱标的Carton No - 使用新的逻辑
            boxes_per_set = int(params.get("盒/套", params.get("boxes_per_set", 1)))
            total_pieces = int(float(data["总张数"]))
            total_sets = math.ceil(math.ceil(total_pieces / pieces_per_box) / boxes_per_set)
            print(f"\n📦 准备计算二级箱标 #{large_box_num} 的Carton No")
            carton_no = split_box_data_processor.calculate_carton_range_for_large_box(
                large_box_num, boxes_per_set, boxes_per_large_box, total_sets
            )
            print(f"📦 二级箱标 #{large_box_num} Carton No计算完成: {carton_no}\n")
            
            # 获取标签模版类型 - 参照常规模版的实现方式
            template_type = params.get("标签模版", "有纸卡备注")
            
            # 绘制箱标表格 - 传入二级包装参数，根据模版类型选择函数，使用计算出的carton_no
            if template_type == "有纸卡备注":
                split_box_renderer.draw_split_box_large_box_table(c, width, height, theme_text, pieces_per_box, 
                                               boxes_per_large_box, 1, serial_range, 
                                               carton_no, remark_text, serial_font_size)
            else:  # "无纸卡备注"
                split_box_renderer.draw_split_box_large_box_table_no_paper_card(c, width, height, theme_text, pieces_per_box, 
                                               boxes_per_large_box, 1, serial_range, 
                                               carton_no, remark_text, serial_font_size)

        c.save()

