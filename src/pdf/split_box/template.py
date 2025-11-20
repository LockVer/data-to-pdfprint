"""
Split Box Template - Multi-level PDF generation with special serial number logic
"""
import math
import re
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
from src.utils.carton_summary_generator import generate_carton_summary_for_template


def _clean_for_filename(text: str) -> str:
    r"""
    清理文本使其适合作为Windows/macOS文件名

    处理问题：
    1. 换行符 (\n, \r) - Excel单元格中的换行会导致Windows文件名错误
    2. Windows非法字符 (< > : " / \ | ? *)
    3. 控制字符（ASCII 0-31）

    Args:
        text: 原始文本

    Returns:
        清理后的安全文本
    """
    if not text:
        return ""

    # 转为字符串并清理
    text = str(text)

    # 1. 替换换行符为空格（Excel单元格中的换行）
    text = text.replace('\n', ' ').replace('\r', ' ')

    # 2. 移除Windows非法字符: < > : " / \ | ? * 和控制字符（ASCII 0-31）
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', text)

    # 3. 移除前后空格和点号（Windows不允许文件名以这些结尾）
    text = text.strip('. ')

    # 4. 压缩多余的空格和下划线
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'_+', '_', text)

    return text


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
        创建有小箱模式的PDF
        """
        # 计算数量 - 三级结构：张→盒→小箱→大箱
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        small_boxes_per_large_box = int(params["小箱/大箱"])

        # 计算各级数量，基于套数计算而非简单除法
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 基于套数的正确计算
        boxes_per_set = int(params.get("盒/套", params.get("boxes_per_set", 1)))
        total_sets = math.ceil(total_boxes / boxes_per_set)
        
        # 计算每套需要的小箱数和大箱数 - 区分能力参数和实际箱数
        small_boxes_per_set_ratio = boxes_per_set / boxes_per_small_box
        actual_small_boxes_per_set = math.ceil(small_boxes_per_set_ratio)
        large_boxes_per_set_ratio = actual_small_boxes_per_set / small_boxes_per_large_box
        actual_large_boxes_per_set = math.ceil(large_boxes_per_set_ratio)
        
        # 基于套数计算总数量
        total_small_boxes = total_sets * actual_small_boxes_per_set
        if large_boxes_per_set_ratio >= 1:
            total_large_boxes = total_sets * actual_large_boxes_per_set
        else:
            # 多套分一个大箱：total_large_boxes = ceil(total_sets / sets_per_large_box)
            sets_per_large_box = math.ceil(1 / large_boxes_per_set_ratio)
            total_large_boxes = math.ceil(total_sets / sets_per_large_box)
        
        print(f"🔍 有小箱模式数量计算:")
        print(f"    总张数: {total_pieces}, 张/盒: {pieces_per_box}")
        print(f"    总盒数: {total_boxes} = ceil({total_pieces} ÷ {pieces_per_box})")
        print(f"    盒/套: {boxes_per_set}, 总套数: {total_sets} = ceil({total_boxes} ÷ {boxes_per_set})")
        print(f"    每套小箱数(能力): {small_boxes_per_set_ratio:.3f}, 实际: {actual_small_boxes_per_set}")
        print(f"    每套大箱数(能力): {large_boxes_per_set_ratio:.3f}, 实际: {actual_large_boxes_per_set}")
        print(f"    总小箱数: {total_small_boxes} = {total_sets} × {actual_small_boxes_per_set}")
        if large_boxes_per_set_ratio >= 1:
            print(f"    总大箱数: {total_large_boxes} = {total_sets} × {actual_large_boxes_per_set}")
        else:
            sets_per_large_box = math.ceil(1 / large_boxes_per_set_ratio)
            print(f"    总大箱数: {total_large_boxes} = ceil({total_sets} ÷ {sets_per_large_box}) (多套分一个大箱)")

        # 创建输出目录 - 新格式：编号+英文名+中文名+标签
        clean_customer_code = _clean_for_filename(data['客户名称编码'])  # 编号
        clean_label_name = _clean_for_filename(data['标签名称'])  # 英文名
        clean_chinese_name = _clean_for_filename(params.get("中文名称", ""))  # 中文名
        folder_name = f"{clean_customer_code}+{clean_label_name}+{clean_chinese_name}+标签"
        full_output_dir = Path(output_dir) / folder_name
        full_output_dir.mkdir(parents=True, exist_ok=True)

        # 获取参数和日期时间戳
        # 清理中文名称（可能包含Excel换行符\n和Windows非法字符）
        chinese_name = clean_chinese_name  # 使用已清理的中文名称
        english_name = clean_label_name  # 使用已清理的标签名称
        # 清理客户编号（可能包含Windows非法字符如冒号）
        customer_code = clean_customer_code  # 使用已清理的客户编码
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        generated_files = {}

        # 检查是否需要生成盒标
        has_box_label = params.get("是否有盒标", False)
        
        if has_box_label:
            # 生成分盒盒标 (分盒模板固定使用外观一，无需用户选择)
            selected_appearance = params["选择外观"]  # 固定为外观一
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
            data, params, str(large_box_path), total_large_boxes, excel_file_path, large_boxes_per_set_ratio
        )
        generated_files["大箱标"] = str(large_box_path)

        # 生成外箱汇总表
        try:
            # 计算每箱盒数（有小箱的情况：盒/小箱 × 小箱/大箱）
            boxes_per_large_box = boxes_per_small_box * small_boxes_per_large_box

            summary_file_path = generate_carton_summary_for_template(
                output_dir=str(full_output_dir),
                data=data,
                params=params,
                total_large_boxes=total_large_boxes,
                boxes_per_large_box=boxes_per_large_box,
                total_boxes=total_boxes
            )
            generated_files["外箱汇总表"] = summary_file_path
            print(f"✅ 外箱汇总表已生成: {summary_file_path}")
        except Exception as e:
            print(f"⚠️ 外箱汇总表生成失败: {e}")
            # 汇总表生成失败不影响主流程

        return generated_files
    
    def _create_two_level_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """
        创建无小箱模式的PDF
        """
        # 计算数量 - 二级结构：张→盒→箱
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        # 参数映射问题：在UI的"无小箱"模式下，"盒/小箱"实际存储的是"盒/大箱"的值
        has_small_box = params.get("是否有小箱", True)
        if has_small_box:
            # 有小箱模式：盒/箱 = 盒/小箱 × 小箱/大箱
            boxes_per_small_box = int(params["盒/小箱"]) 
            small_boxes_per_large_box = int(params["小箱/大箱"])
            boxes_per_large_box = boxes_per_small_box * small_boxes_per_large_box
            print(f"✅ 有小箱模式计算: 盒/大箱 = 盒/小箱({boxes_per_small_box}) × 小箱/大箱({small_boxes_per_large_box}) = {boxes_per_large_box}")
        else:
            # 无小箱模式：直接使用"盒/小箱"中存储的"盒/大箱"值
            boxes_per_large_box = int(params["盒/小箱"])  # 注意：这里存储的实际是"盒/大箱"
            print(f"✅ 无小箱模式计算: 盒/大箱 = {boxes_per_large_box} (直接从UI获取)")

        # 计算各级数量
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 大箱数应该基于套数计算，而不是简单的总盒数除法
        boxes_per_set = int(params.get("盒/套", params.get("boxes_per_set", 1)))
        total_sets = math.ceil(total_boxes / boxes_per_set)
        
        # 区分能力参数和实际箱数
        large_boxes_per_set_ratio = boxes_per_set / boxes_per_large_box
        actual_large_boxes_per_set = math.ceil(large_boxes_per_set_ratio)
        
        if large_boxes_per_set_ratio >= 1:
            total_large_boxes = total_sets * actual_large_boxes_per_set
        else:
            # 多套分一个大箱：total_large_boxes = ceil(total_sets / sets_per_large_box)
            sets_per_large_box = math.ceil(1 / large_boxes_per_set_ratio)
            total_large_boxes = math.ceil(total_sets / sets_per_large_box)
        
        print(f"🔍 无小箱模式数量计算:")
        print(f"    总张数: {total_pieces}")
        print(f"    张/盒: {pieces_per_box}")
        print(f"    总盒数: {total_boxes} = ceil({total_pieces} ÷ {pieces_per_box}) = ceil({total_pieces / pieces_per_box})")
        print(f"    盒/套: {boxes_per_set}")
        print(f"    总套数: {total_sets} = ceil({total_boxes} ÷ {boxes_per_set}) = ceil({total_boxes / boxes_per_set})")
        print(f"    盒/大箱: {boxes_per_large_box}")
        print(f"    每套大箱数(能力): {large_boxes_per_set_ratio:.3f}, 实际: {actual_large_boxes_per_set}")
        if large_boxes_per_set_ratio >= 1:
            print(f"    总大箱数: {total_large_boxes} = {total_sets} × {actual_large_boxes_per_set}")
        else:
            sets_per_large_box = math.ceil(1 / large_boxes_per_set_ratio)
            print(f"    总大箱数: {total_large_boxes} = ceil({total_sets} ÷ {sets_per_large_box}) (多套分一个大箱)")

        # 创建输出目录 - 新格式：编号+英文名+中文名+标签
        clean_customer_code = _clean_for_filename(data['客户名称编码'])  # 编号
        clean_label_name = _clean_for_filename(data['标签名称'])  # 英文名
        clean_chinese_name = _clean_for_filename(params.get("中文名称", ""))  # 中文名
        folder_name = f"{clean_customer_code}+{clean_label_name}+{clean_chinese_name}+标签"
        full_output_dir = Path(output_dir) / folder_name
        full_output_dir.mkdir(parents=True, exist_ok=True)

        # 获取参数和日期时间戳
        # 清理中文名称（可能包含Excel换行符\n和Windows非法字符）
        chinese_name = clean_chinese_name  # 使用已清理的中文名称
        english_name = clean_label_name  # 使用已清理的标签名称
        # 清理客户编号（可能包含Windows非法字符如冒号）
        customer_code = clean_customer_code  # 使用已清理的客户编码
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        generated_files = {}

        # 检查是否需要生成盒标
        has_box_label = params.get("是否有盒标", False)
        
        if has_box_label:
            # 生成分盒盒标 (分盒模板固定使用外观一，无需用户选择)
            selected_appearance = params["选择外观"]  # 固定为外观一
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
            data, params, str(large_box_path), total_large_boxes, total_boxes, boxes_per_large_box, excel_file_path, large_boxes_per_set_ratio
        )
        generated_files["箱标"] = str(large_box_path)

        # 生成外箱汇总表
        try:
            # 无小箱的情况，每箱盒数就是 boxes_per_large_box
            summary_file_path = generate_carton_summary_for_template(
                output_dir=str(full_output_dir),
                data=data,
                params=params,
                total_large_boxes=total_large_boxes,
                boxes_per_large_box=boxes_per_large_box,
                total_boxes=total_boxes
            )
            generated_files["外箱汇总表"] = summary_file_path
            print(f"✅ 外箱汇总表已生成: {summary_file_path}")
        except Exception as e:
            print(f"⚠️ 外箱汇总表生成失败: {e}")
            # 汇总表生成失败不影响主流程

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
        boxes_per_set = int(params["盒/套"])
        print(f"✅ 分盒盒标参数: 盒/套={boxes_per_set}, 盒/小箱={boxes_per_small_box}, 小箱/大箱={small_boxes_per_large_box}")
        
        # 直接创建单个PDF文件，包含所有盒标（移除分页限制）
        self._create_single_split_box_label_file(
            data, params, output_path, style, 
            1, total_boxes, top_text, base_number, boxes_per_set, boxes_per_small_box, small_boxes_per_large_box
        )

    def _create_single_split_box_label_file(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, 
                                           style: str, start_box: int, end_box: int, top_text: str, base_number: str, boxes_per_set: int, boxes_per_small_box: int, small_boxes_per_large_box: int):
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
        # 清理中文名称（可能包含Excel换行符\n和Windows非法字符）
        chinese_name = _clean_for_filename(params.get("中文名称", ""))
        
        # 生成指定范围的盒标
        for box_num in range(start_box, end_box + 1):
            # 🔥 新增：在第一个标签时添加空白首页（外观1和外观2都支持）
            if box_num == start_box and style in ["外观一", "外观二"] and chinese_name:
                print(f"📝 生成分盒盒标空白首页({style}): {chinese_name}")
                if style == "外观一":
                    # 外观一：居中显示的空白首页
                    split_box_renderer.render_blank_first_page(c, width, height, chinese_name)
                else:  # 外观二
                    # 外观二：左对齐显示的空白首页
                    split_box_renderer.render_blank_first_page_appearance_two(c, width, height, chinese_name)
                c.showPage()
                c.setFillColor(cmyk_black)

            if box_num > start_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 使用新的盒标Serial计算方法：父级编号为套，子级编号为盒
            current_number = split_box_data_processor.generate_box_serial_with_set_logic(
                base_number, box_num, boxes_per_set
            )

            # 根据选择的外观渲染
            if style == "外观一":
                split_box_renderer.render_appearance_one(c, width, top_text, current_number, top_text_y, serial_number_y)
            else:  # 外观二
                # 获取票数信息用于外观二
                total_pieces = int(float(data["总张数"]))
                pieces_per_box = int(params["张/盒"])
                split_box_renderer.render_appearance_two(c, width, self.page_size, top_text, pieces_per_box, current_number, top_text_y, serial_number_y)

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
        
        # 从remainder_info获取total_boxes
        total_boxes = remainder_info.get("total_boxes", 0)
        
        # 直接创建单个PDF文件，包含所有小箱标
        self._create_single_split_box_small_box_label_file(
            data, params, output_path, 1, total_small_boxes,
            theme_text, base_number, remark_text, pieces_per_box, 
            boxes_per_set, boxes_per_small_box, total_small_boxes, small_boxes_per_large_box, total_boxes, serial_font_size
        )

    def _create_single_split_box_small_box_label_file(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
                                                 start_small_box: int, end_small_box: int, theme_text: str, base_number: str,
                                                 remark_text: str, pieces_per_box: int, boxes_per_set: int, boxes_per_small_box: int, 
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
            # 清理中文名称（可能包含Excel换行符\n和Windows非法字符）
            chinese_name = _clean_for_filename(params.get("中文名称", ""))
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

            # 🔧 根据模式选择Serial生成逻辑
            if boxes_per_set > 1:  # 分/套盒模式
                serial_range = split_box_data_processor.generate_set_based_small_box_serial_range(
                    small_box_num, base_number, boxes_per_set, boxes_per_small_box, total_boxes
                )
            else:  # 传统分盒模式
                serial_range = split_box_data_processor.generate_split_small_box_serial_range(
                    base_number, small_box_num, boxes_per_small_box, small_boxes_per_large_box, total_boxes
                )

            # 🔧 使用新的quantity计算方法
            actual_pieces_in_small_box = split_box_data_processor.calculate_actual_quantity_for_small_box(
                small_box_num, pieces_per_box, boxes_per_small_box, total_boxes
            )

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
                                     total_large_boxes: int, excel_file_path: str = None, large_boxes_per_set_ratio: float = None):
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
        
        # 计算large_boxes_per_set_ratio（如果没有传入）
        if large_boxes_per_set_ratio is None:
            boxes_per_large_box = boxes_per_small_box * small_boxes_per_large_box
            large_boxes_per_set_ratio = boxes_per_set / boxes_per_large_box
        
        # 直接创建单个PDF文件，包含所有大箱标
        self._create_single_split_box_large_box_label_file(
            data, params, output_path, 1, total_large_boxes,
            theme_text, base_number, remark_text, pieces_per_box, 
            boxes_per_set, boxes_per_small_box, small_boxes_per_large_box, total_large_boxes, total_sets, serial_font_size, large_boxes_per_set_ratio
        )

    def _create_single_split_box_large_box_label_file(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
                                                 start_large_box: int, end_large_box: int, theme_text: str, base_number: str,
                                                 remark_text: str, pieces_per_box: int, boxes_per_set: int, boxes_per_small_box: int, 
                                                 small_boxes_per_large_box: int, total_large_boxes: int, total_sets: int, serial_font_size: int = 10, large_boxes_per_set_ratio: float = None):
        """创建单个分盒大箱标PDF文件 - 完全参考小箱标"""
        
        # 计算large_boxes_per_set_ratio（如果没有传入）
        if large_boxes_per_set_ratio is None:
            boxes_per_large_box = boxes_per_small_box * small_boxes_per_large_box
            large_boxes_per_set_ratio = boxes_per_set / boxes_per_large_box
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
            # 清理中文名称（可能包含Excel换行符\n和Windows非法字符）
            chinese_name = _clean_for_filename(params.get("中文名称", ""))
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
            
            # 🔧 根据模式选择Serial生成逻辑
            if boxes_per_set > 1:  # 分/套盒模式
                serial_range = split_box_data_processor.generate_set_based_large_box_serial_range(
                    large_box_num, base_number, boxes_per_set, boxes_per_small_box, small_boxes_per_large_box, total_boxes
                )
            else:  # 传统分盒模式
                serial_range = split_box_data_processor.generate_split_large_box_serial_range(
                    base_number, large_box_num, small_boxes_per_large_box, boxes_per_small_box, total_boxes
                )
            
            # 计算大箱标的Carton No - 基于最新逻辑整理  
            boxes_per_large_box = boxes_per_small_box * small_boxes_per_large_box
            print(f"\n📦 准备计算大箱标 #{large_box_num} 的Carton No")
            carton_no = split_box_data_processor.calculate_carton_range_for_large_box(
                large_box_num, large_boxes_per_set_ratio, total_sets
            )
            print(f"📦 大箱标 #{large_box_num} Carton No计算完成: {carton_no}\n")
            
            # 🔧 使用新的quantity计算方法
            actual_quantity_for_large_box = split_box_data_processor.calculate_actual_quantity_for_large_box(
                large_box_num, pieces_per_box, boxes_per_small_box, small_boxes_per_large_box, total_boxes, boxes_per_set
            )
            
            # 获取标签模版类型 - 参照常规模版的实现方式
            template_type = params.get("标签模版", "有纸卡备注")
            
            # 绘制大箱标表格 - 使用预计算的quantity值，根据模版类型选择函数
            if template_type == "有纸卡备注":
                split_box_renderer.draw_split_box_large_box_table(c, width, height, theme_text, actual_quantity_for_large_box,
                                               serial_range, carton_no, remark_text, serial_font_size)
            else:  # "无纸卡备注"
                split_box_renderer.draw_split_box_large_box_table_no_paper_card(c, width, height, theme_text, actual_quantity_for_large_box,
                                               serial_range, carton_no, remark_text, serial_font_size)

        c.save()

    def _create_two_level_large_box_label(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, 
                                     total_large_boxes: int, total_boxes: int, boxes_per_large_box: int, excel_file_path: str = None, large_boxes_per_set_ratio: float = None):
        """创建无小箱模式的箱标"""
        # 获取Excel数据 - 使用关键字提取，与大箱标相同
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用统一数据处理后的标准四字段（优先使用传入的data参数）
        theme_text = data.get('标签名称') or 'Unknown Title'
        base_number = data.get('开始号') or 'DEFAULT01001'
        remark_text = data.get('客户名称编码') or 'Unknown Client'
        print(f"✅ 分盒箱标使用统一数据: 主题='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}'")
        
        # 计算参数 - 箱标专用（无小箱模式）
        pieces_per_box = int(params["张/盒"])  # 第一个参数：张/盒
        serial_font_size = int(params.get("序列号字体大小", 10))
        print(f"✅ 分盒箱标参数: 盒/箱={boxes_per_large_box}, 序列号字体大小={serial_font_size}")
        
        # 计算large_boxes_per_set_ratio参数
        boxes_per_set = int(params.get("盒/套", params.get("boxes_per_set", 1)))
        large_boxes_per_set_ratio = boxes_per_set / boxes_per_large_box
        
        # 直接创建单个PDF文件，包含所有箱标
        self._create_single_two_level_large_box_label_file(
            data, params, output_path, 1, total_large_boxes,
            theme_text, base_number, remark_text, pieces_per_box, 
            boxes_per_large_box, total_large_boxes, total_boxes, serial_font_size, large_boxes_per_set_ratio
        )

    def _create_single_two_level_large_box_label_file(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
                                                 start_large_box: int, end_large_box: int, theme_text: str, base_number: str,
                                                 remark_text: str, pieces_per_box: int, boxes_per_large_box: int, 
                                                 total_large_boxes: int, total_boxes: int, serial_font_size: int = 10, large_boxes_per_set_ratio: float = None):
        """创建单个分盒箱标PDF文件（无小箱模式）"""
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
            # 清理中文名称（可能包含Excel换行符\n和Windows非法字符）
            chinese_name = _clean_for_filename(params.get("中文名称", ""))
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

            # 计算无小箱模式箱标的Carton No - 使用新的逻辑
            boxes_per_set = int(params.get("盒/套", params.get("boxes_per_set", 1)))
            
            # 🔧 根据模式选择Serial生成逻辑（无小箱模式）
            if boxes_per_set > 1:  # 分/套盒模式
                serial_range = split_box_data_processor.generate_set_based_large_box_serial_range(
                    large_box_num, base_number, boxes_per_set, boxes_per_large_box, 1, total_boxes
                )
            else:  # 传统分盒模式
                # 无小箱模式：复用大箱标逻辑，但设置 small_boxes_per_large_box = 1，进位阈值 = boxes_per_large_box
                serial_range = split_box_data_processor.generate_split_large_box_serial_range(
                    base_number, large_box_num, 1, boxes_per_large_box, total_boxes
                )
            total_pieces = int(float(data["总张数"]))
            total_sets = math.ceil(math.ceil(total_pieces / pieces_per_box) / boxes_per_set)
            
            # 计算large_boxes_per_set_ratio（如果没有传入）
            if large_boxes_per_set_ratio is None:
                large_boxes_per_set_ratio = boxes_per_set / boxes_per_large_box
            
            print(f"\n📦 准备计算无小箱模式箱标 #{large_box_num} 的Carton No")
            carton_no = split_box_data_processor.calculate_carton_range_for_large_box(
                large_box_num, large_boxes_per_set_ratio, total_sets
            )
            print(f"📦 无小箱模式箱标 #{large_box_num} Carton No计算完成: {carton_no}\n")
            
            # 🔧 使用新的quantity计算方法（二级模式：盒直接装到大箱）
            total_boxes = math.ceil(total_pieces / pieces_per_box)
            boxes_per_set = int(params.get("盒/套", params.get("boxes_per_set", 1)))
            actual_quantity_for_large_box = split_box_data_processor.calculate_actual_quantity_for_large_box(
                large_box_num, pieces_per_box, boxes_per_large_box, 1, total_boxes, boxes_per_set
            )
            
            # 获取标签模版类型 - 参照常规模版的实现方式
            template_type = params.get("标签模版", "有纸卡备注")
            
            # 绘制箱标表格 - 使用预计算的quantity值，根据模版类型选择函数
            if template_type == "有纸卡备注":
                split_box_renderer.draw_split_box_large_box_table(c, width, height, theme_text, actual_quantity_for_large_box,
                                               serial_range, carton_no, remark_text, serial_font_size)
            else:  # "无纸卡备注"
                split_box_renderer.draw_split_box_large_box_table_no_paper_card(c, width, height, theme_text, actual_quantity_for_large_box,
                                               serial_range, carton_no, remark_text, serial_font_size)

        c.save()

