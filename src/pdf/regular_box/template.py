"""
常规模板 - 标准的多级标签PDF生成
"""
import math
import re
from datetime import datetime
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

# 导入常规模板专属数据处理器和渲染器
from src.pdf.regular_box.data_processor import regular_data_processor
from src.pdf.regular_box.renderer import regular_renderer
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


class RegularTemplate(PDFBaseUtils):
    """常规模板处理类"""
    
    def __init__(self):
        """初始化常规模板"""
        super().__init__()
    
    def create_multi_level_pdfs(self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None) -> Dict[str, str]:
        """
        创建常规模板的多级标签PDF

        Args:
            data: Excel数据
            params: 用户参数 (张/盒, 盒/小箱, 小箱/大箱, 选择外观)
            output_dir: 输出目录

        Returns:
            生成的文件路径字典
        """
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

        # 计算各级数量
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)
        total_large_boxes = math.ceil(total_small_boxes / small_boxes_per_large_box)

        # 计算余数信息
        remaining_pieces_in_last_box = total_pieces % pieces_per_box
        remaining_boxes_in_last_small_box = total_boxes % boxes_per_small_box
        remaining_small_boxes_in_last_large_box = total_small_boxes % small_boxes_per_large_box

        remainder_info = {
            "total_boxes": total_boxes,
            "remaining_pieces_in_last_box": (
                pieces_per_box if remaining_pieces_in_last_box == 0 else remaining_pieces_in_last_box
            ),
            "remaining_boxes_in_last_small_box": (
                boxes_per_small_box if remaining_boxes_in_last_small_box == 0 else remaining_boxes_in_last_small_box
            ),
            "remaining_small_boxes_in_last_large_box": (
                small_boxes_per_large_box if remaining_small_boxes_in_last_large_box == 0 else remaining_small_boxes_in_last_large_box
            ),
        }

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
            # 生成盒标 (只生成用户选择的外观)
            selected_appearance = params["选择外观"]
            # 文件名格式：客户编号_中文名称_英文名称_盒标_外观类型_日期时间戳
            box_label_filename = f"{customer_code}_{chinese_name}_{english_name}_盒标_{selected_appearance}_{timestamp}.pdf"
            box_label_path = full_output_dir / box_label_filename

            self._create_box_label(data, params, str(box_label_path), selected_appearance, excel_file_path)
            generated_files["盒标"] = str(box_label_path)
        else:
            print("⏭️ 用户选择无盒标，跳过盒标生成")

        # 生成小箱标
        # 文件名格式：客户编号_中文名称_英文名称_小箱标_日期时间戳
        small_box_filename = f"{customer_code}_{chinese_name}_{english_name}_小箱标_{timestamp}.pdf"
        small_box_path = full_output_dir / small_box_filename
        self._create_small_box_label(
            data, params, str(small_box_path), total_small_boxes, remainder_info, total_boxes, excel_file_path
        )
        generated_files["小箱标"] = str(small_box_path)

        # 生成大箱标
        # 文件名格式：客户编号_中文名称_英文名称_大箱标_日期时间戳
        large_box_filename = f"{customer_code}_{chinese_name}_{english_name}_大箱标_{timestamp}.pdf"
        large_box_path = full_output_dir / large_box_filename
        self._create_large_box_label(
            data, params, str(large_box_path), total_large_boxes, total_small_boxes, remainder_info, total_boxes, excel_file_path
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
        创建二级包装的PDF（无小箱）
        """
        # 计算数量 - 二级结构：张→盒→箱
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        boxes_per_large_box = int(params["盒/小箱"])  # 在二级模式下，这实际上是盒/箱

        # 计算各级数量
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        total_large_boxes = math.ceil(total_boxes / boxes_per_large_box)

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
            # 生成盒标 (只生成用户选择的外观)
            selected_appearance = params["选择外观"]
            # 文件名格式：客户编号_中文名称_英文名称_盒标_外观类型_日期时间戳
            box_label_filename = f"{customer_code}_{chinese_name}_{english_name}_盒标_{selected_appearance}_{timestamp}.pdf"
            box_label_path = full_output_dir / box_label_filename

            self._create_box_label(data, params, str(box_label_path), selected_appearance, excel_file_path)
            generated_files["盒标"] = str(box_label_path)
        else:
            print("⏭️ 用户选择无盒标，跳过盒标生成")

        # 生成箱标（复用大箱标逻辑但文件名为箱标）
        # 文件名格式：客户编号_中文名称_英文名称_箱标_日期时间戳
        large_box_filename = f"{customer_code}_{chinese_name}_{english_name}_箱标_{timestamp}.pdf"
        large_box_path = full_output_dir / large_box_filename
        
        # 构造虚拟参数来复用大箱标逻辑
        virtual_params = params.copy()
        virtual_params["小箱/大箱"] = 1  # 设置为1表示跳过小箱层级
        
        self._create_two_level_large_box_label(
            data, virtual_params, str(large_box_path), total_large_boxes, total_boxes, boxes_per_large_box, excel_file_path
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

    def _create_box_label(self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, style: str, excel_file_path: str = None):
        """创建盒标 - 支持分页限制的多页PDF"""
        # 计算总盒数
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 使用统一数据处理后的标准四字段（优先使用传入的data参数）
        top_text = data.get('标签名称') or 'Unknown Title'
        base_number = data.get('开始号') or 'DSK00001'
        print(f"✅ 常规盒标使用统一数据: 主题='{top_text}', 开始号='{base_number}'")
        
        # 生成单个PDF文件（不分页）
        self._create_single_box_label_file(
            data, params, output_path, style,
            1, total_boxes, top_text, base_number
        )

    def _create_single_box_label_file(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, 
        style: str, start_box: int, end_box: int, top_text: str, base_number: str
    ):
        """创建单个盒标PDF文件"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"盒标-{style}-{start_box}到{end_box}")
        c.setSubject("Box Label")
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
                print(f"📝 生成常规盒标空白首页({style}): {chinese_name}")
                if style == "外观一":
                    # 外观一：居中显示的空白首页
                    regular_renderer.render_blank_first_page(c, width, height, chinese_name)
                else:  # 外观二
                    # 外观二：左对齐显示的空白首页
                    regular_renderer.render_blank_first_page_appearance_two(c, width, height, chinese_name)
                c.showPage()
                c.setFillColor(cmyk_black)
            
            if box_num > start_box:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 解析基础序列号格式
            import re
            # 修复: 使用只匹配字母的前缀，避免将数字当作前缀的一部分
            # 例如 'CCD01001' 应解析为 prefix='CCD', number='01001'
            match = re.search(r'([A-Za-z]+)(-?)(\d+)', base_number)
            if match:
                # 获取前缀、连字符和数字
                prefix = match.group(1)
                separator = match.group(2)  # 保留连字符（如果存在）
                number_str = match.group(3)
                base_num = int(number_str)
                digits = len(number_str)

                # 计算当前序列号
                current_num = base_num + (box_num - 1)
                # 保持与Excel中数字部分相同的位数
                current_number = f"{prefix}{separator}{current_num:0{digits}d}"
            else:
                # 如果无法解析，使用简单递增，保持5位数字格式（向后兼容）
                current_number = f"BOX{box_num:05d}"

            # 根据选择的外观渲染
            if style == "外观一":
                regular_renderer.render_appearance_one(c, width, top_text, current_number, top_text_y, serial_number_y)
            else:
                # 获取票数信息用于外观二
                total_pieces = int(float(data["总张数"]))
                pieces_per_box = int(params["张/盒"])
                regular_renderer.render_appearance_two(c, width, self.page_size, top_text, pieces_per_box, current_number, top_text_y, serial_number_y)

        c.save()



    def _create_small_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        total_small_boxes: int,
        remainder_info: Dict[str, Any],
        total_boxes: int,
        excel_file_path: str = None,
    ):
        """创建小箱标"""
        # 获取Excel数据 - 使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用统一数据处理后的标准四字段（优先使用传入的data参数）
        theme_text = data.get('标签名称') or 'Unknown Title'
        base_number = data.get('开始号') or 'DEFAULT01001'
        remark_text = data.get('客户名称编码') or 'Unknown Client'
        print(f"✅ 常规小箱标使用统一数据: 主题='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}'")
        
        # 计算参数
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        pieces_per_small_box = pieces_per_box * boxes_per_small_box
        serial_font_size = int(params.get("序列号字体大小", 10))
        
        # 生成单个PDF文件（不分页）
        self._create_single_small_box_label_file(
            data, params, output_path,
            1, total_small_boxes, theme_text, base_number, remark_text, 
            pieces_per_small_box, boxes_per_small_box, total_small_boxes, total_boxes, serial_font_size
        )

    def _create_single_small_box_label_file(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
        start_small_box: int, end_small_box: int, theme_text: str, base_number: str,
        remark_text: str, pieces_per_small_box: int, boxes_per_small_box: int, 
        total_small_boxes: int, total_boxes: int, serial_font_size: int = 10
    ):
        """创建单个小箱标PDF文件"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"小箱标-{start_small_box}到{end_small_box}")
        c.setSubject("Small Box Label")
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
                regular_renderer.render_empty_box_label(c, width, height, chinese_name, remark_text)
            else:  # "无纸卡备注"
                regular_renderer.render_empty_box_label_no_paper_card(c, width, height, chinese_name, remark_text)
            
            c.showPage()
            c.setFillColor(cmyk_black)

        # 生成指定范围的小箱标
        for small_box_num in range(start_small_box, end_small_box + 1):
            if small_box_num > start_small_box or start_small_box == 1:  # 修改条件，考虑空标签页
                if not (small_box_num == start_small_box and start_small_box == 1):  # 避免重复showPage
                    c.showPage()
                    c.setFillColor(cmyk_black)

            # 🔧 使用修复后的数据处理器计算序列号范围（包含边界检查）
            serial_range = regular_data_processor.generate_regular_small_box_serial_range(
                base_number, small_box_num, boxes_per_small_box, total_boxes
            )

            # 🔧 计算当前小箱的实际张数（考虑最后一小箱的边界情况）
            pieces_per_box = int(params["张/盒"])
            # 计算当前小箱实际包含的盒数
            start_box = (small_box_num - 1) * boxes_per_small_box + 1
            end_box = min(start_box + boxes_per_small_box - 1, total_boxes)
            actual_boxes_in_small_box = end_box - start_box + 1
            actual_pieces_in_small_box = actual_boxes_in_small_box * pieces_per_box

            # 计算小箱标Carton No - 格式：当前小箱/总小箱数
            carton_no = regular_data_processor.calculate_carton_number_for_small_box(small_box_num, total_small_boxes)
            
            # 获取标签模版类型
            template_type = params.get("标签模版", "有纸卡备注")
            
            # 绘制小箱标表格（使用实际张数，传入模版类型和序列号字体大小）
            regular_renderer.draw_small_box_table(c, width, height, theme_text, actual_pieces_in_small_box, 
                                                 serial_range, carton_no, remark_text, template_type, serial_font_size)

        c.save()


    def _create_large_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        total_large_boxes: int,
        total_small_boxes: int,
        remainder_info: Dict[str, Any],
        total_boxes: int,
        excel_file_path: str = None,
    ):
        """创建大箱标"""
        # 获取Excel数据 - 使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用统一数据处理后的标准四字段（优先使用传入的data参数）
        theme_text = data.get('标签名称') or 'Unknown Title'
        base_number = data.get('开始号') or 'DEFAULT01001'
        remark_text = data.get('客户名称编码') or 'Unknown Client'
        print(f"✅ 常规大箱标使用统一数据: 主题='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}'")
        
        # 计算参数 - 大箱标专用
        pieces_per_box = int(params["张/盒"])  
        boxes_per_small_box = int(params["盒/小箱"]) 
        small_boxes_per_large_box = int(params["小箱/大箱"])  
        
        pieces_per_large_box = pieces_per_box * boxes_per_small_box * small_boxes_per_large_box
        
        # 获取序列号字体大小参数
        serial_font_size = int(params.get("序列号字体大小", 10))
        
        # 生成单个PDF文件（不分页）
        self._create_single_large_box_label_file(
            data, params, output_path,
            1, total_large_boxes,
            theme_text, base_number, remark_text, pieces_per_large_box, 
            boxes_per_small_box, small_boxes_per_large_box, total_large_boxes, total_boxes, serial_font_size
        )

    def _create_single_large_box_label_file(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
        start_large_box: int, end_large_box: int, theme_text: str, base_number: str,
        remark_text: str, pieces_per_large_box: int, boxes_per_small_box: int, 
        small_boxes_per_large_box: int, total_large_boxes: int, total_boxes: int, serial_font_size: int = 10
    ):
        """创建单个大箱标PDF文件"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"大箱标-{start_large_box}到{end_large_box}")
        c.setSubject("Large Box Label")
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
                regular_renderer.render_empty_box_label(c, width, height, chinese_name, remark_text)
            else:  # "无纸卡备注"
                regular_renderer.render_empty_box_label_no_paper_card(c, width, height, chinese_name, remark_text)
            
            c.showPage()
            c.setFillColor(cmyk_black)

        # 生成指定范围的大箱标
        for large_box_num in range(start_large_box, end_large_box + 1):
            if large_box_num > start_large_box or start_large_box == 1:  # 修改条件，考虑空标签页
                if not (large_box_num == start_large_box and start_large_box == 1):  # 避免重复showPage
                    c.showPage()
                    c.setFillColor(cmyk_black)

            # 🔧 使用修复后的数据处理器计算序列号范围（包含边界检查）
            serial_range = regular_data_processor.generate_regular_large_box_serial_range(
                base_number, large_box_num, small_boxes_per_large_box, boxes_per_small_box, total_boxes
            )

            # 🔧 计算当前大箱的实际张数（考虑最后一大箱的边界情况）
            pieces_per_box = int(params["张/盒"])
            # 计算当前大箱实际包含的盒数
            boxes_per_large_box = boxes_per_small_box * small_boxes_per_large_box
            start_box = (large_box_num - 1) * boxes_per_large_box + 1
            end_box = min(start_box + boxes_per_large_box - 1, total_boxes)
            actual_boxes_in_large_box = end_box - start_box + 1
            actual_pieces_in_large_box = actual_boxes_in_large_box * pieces_per_box

            # 计算大箱标Carton No - 格式：当前大箱/总大箱数  
            carton_no = regular_data_processor.calculate_carton_range_for_large_box(large_box_num, total_large_boxes)
            
            # 获取标签模版类型
            template_type = params.get("标签模版", "有纸卡备注")
            
            # 绘制大箱标表格（使用实际张数，传入模版类型和序列号字体大小）
            regular_renderer.draw_large_box_table(c, width, height, theme_text, actual_pieces_in_large_box,
                                                 serial_range, carton_no, remark_text, template_type, serial_font_size)

        c.save()
    
    def _create_two_level_large_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        total_large_boxes: int,
        total_boxes: int,
        boxes_per_large_box: int,
        excel_file_path: str = None,
    ):
        """创建二级模式的箱标（无小箱）"""
        # 获取Excel数据 - 使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        # 使用统一数据处理后的标准四字段（优先使用传入的data参数）
        theme_text = data.get('标签名称') or 'Unknown Title'
        base_number = data.get('开始号') or 'DEFAULT01001'
        remark_text = data.get('客户名称编码') or 'Unknown Client'
        print(f"✅ 常规箱标使用统一数据: 主题='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}'")
        
        # 计算参数 - 箱标专用（二级模式）
        pieces_per_box = int(params["张/盒"])  
        pieces_per_large_box = pieces_per_box * boxes_per_large_box
        
        # 获取序列号字体大小参数
        serial_font_size = int(params.get("序列号字体大小", 10))
        
        # 生成单个PDF文件（不分页）
        self._create_single_two_level_large_box_label_file(
            data, params, output_path,
            1, total_large_boxes,
            theme_text, base_number, remark_text, pieces_per_large_box, 
            boxes_per_large_box, total_large_boxes, total_boxes, serial_font_size
        )

    def _create_single_two_level_large_box_label_file(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str,
        start_large_box: int, end_large_box: int, theme_text: str, base_number: str,
        remark_text: str, pieces_per_large_box: int, boxes_per_large_box: int, 
        total_large_boxes: int, total_boxes: int, serial_font_size: int = 10
    ):
        """创建单个二级箱标PDF文件"""
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"箱标-{start_large_box}到{end_large_box}")
        c.setSubject("Box Label (Two Level)")
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
                regular_renderer.render_empty_box_label(c, width, height, chinese_name, remark_text)
            else:  # "无纸卡备注"
                regular_renderer.render_empty_box_label_no_paper_card(c, width, height, chinese_name, remark_text)
            
            c.showPage()
            c.setFillColor(cmyk_black)

        # 生成指定范围的箱标
        for large_box_num in range(start_large_box, end_large_box + 1):
            if large_box_num > start_large_box or start_large_box == 1:  # 修改条件，考虑空标签页
                if not (large_box_num == start_large_box and start_large_box == 1):  # 避免重复showPage
                    c.showPage()
                    c.setFillColor(cmyk_black)

            # 🔧 使用修复后的数据处理器计算序列号范围（包含边界检查）
            # 二级模式：复用大箱标逻辑，但设置 small_boxes_per_large_box = 1
            serial_range = regular_data_processor.generate_regular_large_box_serial_range(
                base_number, large_box_num, 1, boxes_per_large_box, total_boxes
            )

            # 🔧 计算当前箱的实际张数（考虑最后一箱的边界情况）
            pieces_per_box = int(params["张/盒"])
            # 计算当前箱实际包含的盒数
            start_box = (large_box_num - 1) * boxes_per_large_box + 1
            end_box = min(start_box + boxes_per_large_box - 1, total_boxes)
            actual_boxes_in_large_box = end_box - start_box + 1
            actual_pieces_in_large_box = actual_boxes_in_large_box * pieces_per_box

            # 计算箱标Carton No - 格式：当前箱/总箱数  
            carton_no = regular_data_processor.calculate_carton_range_for_large_box(large_box_num, total_large_boxes)
            
            # 获取标签模版类型
            template_type = params.get("标签模版", "有纸卡备注")
            
            # 绘制箱标表格（使用实际张数，传入模版类型和序列号字体大小）
            regular_renderer.draw_large_box_table(c, width, height, theme_text, actual_pieces_in_large_box,
                                                 serial_range, carton_no, remark_text, template_type, serial_font_size)

        c.save()
