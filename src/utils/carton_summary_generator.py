"""
外箱装箱信息汇总表生成器
在标签生成完成后，自动生成外箱汇总Excel表格
"""

import pandas as pd
from pathlib import Path
import re
from typing import Dict, Any


def _clean_for_filename(text: str) -> str:
    """
    清理文本使其适合作为Windows/macOS文件名

    Args:
        text: 原始文本

    Returns:
        清理后的安全文本
    """
    if not text:
        return ""

    # 转为字符串并清理
    text = str(text)

    # 1. 替换换行符为空格
    text = text.replace('\n', ' ').replace('\r', ' ')

    # 2. 移除Windows非法字符: < > : " / \ | ? * 和控制字符
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', text)

    # 3. 移除前后空格和点号
    text = text.strip('. ')

    # 4. 压缩多余的空格和下划线
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'_+', '_', text)

    return text


class CartonSummaryGenerator:
    """外箱装箱信息汇总表生成器"""

    def __init__(self):
        """初始化生成器"""
        pass

    def generate_summary(
        self,
        output_dir: str,
        product_code: str,
        chinese_name: str,
        english_name: str,
        pieces_per_box: int,
        total_large_boxes: int,
        boxes_per_large_box: int
    ) -> str:
        """
        生成外箱装箱信息汇总表

        Args:
            output_dir: 输出目录（产品文件夹路径）
            product_code: 产品编号
            chinese_name: 中文名称
            english_name: 英文名称（标签名称）
            pieces_per_box: 每盒数量（张/盒）
            total_large_boxes: 总箱数（总外箱数）
            boxes_per_large_box: 每箱盒数（每个大箱包含的盒数）

        Returns:
            生成的Excel文件路径
        """
        # 清理名称用于显示和文件名
        clean_chinese = _clean_for_filename(chinese_name) if chinese_name else ""
        clean_english = _clean_for_filename(english_name) if english_name else ""
        clean_code = _clean_for_filename(product_code) if product_code else ""

        # 组合显示名称：中文名 + 空格 + 英文名
        display_name = f"{clean_chinese} {clean_english}".strip()
        if not display_name:
            display_name = "未知产品"

        # 创建数据字典
        summary_data = {
            "名称": [display_name],
            "每盒数量": [pieces_per_box],
            "总箱数": [total_large_boxes],
            "每箱盒数": [boxes_per_large_box]
        }

        # 创建DataFrame
        df = pd.DataFrame(summary_data)

        # 生成文件名：外箱汇总-{产品编号}-{英文名}.xlsx
        filename = f"外箱汇总-{clean_code}-{clean_english}.xlsx"
        file_path = Path(output_dir) / filename

        # 保存Excel文件
        try:
            # 使用openpyxl引擎并设置格式
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='外箱汇总')

                # 获取工作表并设置列宽
                worksheet = writer.sheets['外箱汇总']

                # 设置列宽（根据内容调整）
                worksheet.column_dimensions['A'].width = 30  # 名称列
                worksheet.column_dimensions['B'].width = 15  # 每盒数量
                worksheet.column_dimensions['C'].width = 15  # 总箱数
                worksheet.column_dimensions['D'].width = 15  # 每箱盒数

            print(f"✅ 外箱汇总表已生成: {file_path}")
            return str(file_path)

        except Exception as e:
            print(f"❌ 生成外箱汇总表失败: {e}")
            raise


# 创建全局实例
carton_summary_generator = CartonSummaryGenerator()


def generate_carton_summary_for_template(
    output_dir: str,
    data: Dict[str, Any],
    params: Dict[str, Any],
    total_large_boxes: int,
    boxes_per_large_box: int
) -> str:
    """
    为模板生成外箱汇总表的便捷函数

    Args:
        output_dir: 输出目录（产品文件夹路径）
        data: 产品数据字典（包含客户名称编码、标签名称等）
        params: 参数字典（包含中文名称、张/盒等）
        total_large_boxes: 总箱数（总外箱数）
        boxes_per_large_box: 每箱盒数

    Returns:
        生成的Excel文件路径
    """
    # 提取数据
    product_code = data.get('客户名称编码', '')
    english_name = data.get('标签名称', '')
    chinese_name = params.get('中文名称', '')
    pieces_per_box = int(params.get('张/盒', 0))

    # 调用生成器
    return carton_summary_generator.generate_summary(
        output_dir=output_dir,
        product_code=product_code,
        chinese_name=chinese_name,
        english_name=english_name,
        pieces_per_box=pieces_per_box,
        total_large_boxes=total_large_boxes,
        boxes_per_large_box=boxes_per_large_box
    )
