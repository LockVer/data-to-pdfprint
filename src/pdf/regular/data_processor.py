"""
常规模板专属数据处理器
封装常规模板的所有数据提取和处理逻辑，确保与原有功能完全一致
"""

import pandas as pd
import re
import math
from typing import Dict, Any

# 导入现有的通用Excel工具，确保功能一致性
from src.utils.excel_data_extractor import ExcelDataExtractor


class RegularDataProcessor:
    """常规模板专属数据处理器 - 封装现有逻辑，确保功能完全一致"""
    
    def __init__(self):
        """初始化常规数据处理器"""
        self.template_type = "regular"
    
    def extract_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取常规盒标所需的数据 - 使用统一的公共数据提取方法
        """
        # 使用统一的公共数据提取方法
        extractor = ExcelDataExtractor(excel_file_path)
        common_data = extractor.extract_common_data()
        
        return {
            '标签名称': common_data.get('标签名称', 'Unknown Title'),
            '开始号': common_data.get('开始号', 'DSK00001'),
            '客户编码': common_data.get('客户编码', 'Unknown Client')
        }
    
    def extract_small_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取常规小箱标所需的数据 - 使用统一的公共数据提取方法
        """
        # 使用统一的公共数据提取方法
        extractor = ExcelDataExtractor(excel_file_path)
        return extractor.extract_common_data()
    
    def extract_large_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取常规大箱标所需的数据 - 使用统一的公共数据提取方法
        """
        # 使用统一的公共数据提取方法
        extractor = ExcelDataExtractor(excel_file_path)
        return extractor.extract_common_data()
    
    def parse_serial_number_format(self, serial_number: str) -> Dict[str, Any]:
        """
        解析序列号格式 - 与原有逻辑完全一致
        常规模板使用简单的线性递增逻辑
        """
        # 常规模板的序列号格式相对简单（与原代码一致）
        if not serial_number:
            return {
                'prefix': 'DSK',
                'start_number': 1,
                'digits': 5
            }
        
        # 尝试解析DSK00001这种格式
        match = re.search(r'([A-Z]+)(\d+)', serial_number)
        if match:
            prefix = match.group(1)
            start_number = int(match.group(2))
            digits = len(match.group(2))
            
            return {
                'prefix': prefix,
                'start_number': start_number,
                'digits': digits
            }
        else:
            # 如果无法解析，使用默认格式
            return {
                'prefix': 'DSK',
                'start_number': 1,
                'digits': 5
            }
    
    def calculate_quantities(self, total_pieces: int, pieces_per_box: int, 
                           boxes_per_small_box: int, small_boxes_per_large_box: int) -> Dict[str, int]:
        """
        计算各级数量 - 与原有逻辑完全一致
        对应原来各个方法中的数量计算逻辑
        """
        # 计算各级数量
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)
        total_large_boxes = math.ceil(total_small_boxes / small_boxes_per_large_box)
        
        # 计算余数信息（常规模板特有）
        remaining_pieces_in_last_box = total_pieces % pieces_per_box
        remaining_boxes_in_last_small_box = total_boxes % boxes_per_small_box
        remaining_small_boxes_in_last_large_box = total_small_boxes % small_boxes_per_large_box
        
        return {
            'total_pieces': total_pieces,
            'total_boxes': total_boxes,
            'total_small_boxes': total_small_boxes,
            'total_large_boxes': total_large_boxes,
            'pieces_per_box': pieces_per_box,
            'boxes_per_small_box': boxes_per_small_box,
            'small_boxes_per_large_box': small_boxes_per_large_box,
            'remaining_pieces_in_last_box': remaining_pieces_in_last_box,
            'remaining_boxes_in_last_small_box': remaining_boxes_in_last_small_box,
            'remaining_small_boxes_in_last_large_box': remaining_small_boxes_in_last_large_box
        }
    
    def generate_regular_box_serial_number(self, base_number: str, box_num: int) -> str:
        """
        生成常规盒标的序列号 - 与原有逻辑完全一致
        对应原来 _create_regular_box_label 中的序列号生成逻辑
        
        常规模板使用简单的线性递增
        """
        serial_info = self.parse_serial_number_format(base_number)
        
        # 常规模板序列号生成逻辑：简单的线性递增（与原代码完全一致）
        current_number = serial_info['start_number'] + (box_num - 1)
        formatted_number = f"{serial_info['prefix']}{current_number:0{serial_info['digits']}d}"
        
        print(f"📝 常规盒标 #{box_num}: {formatted_number}")
        return formatted_number
    
    def generate_regular_small_box_serial_range(self, base_number: str, small_box_num: int, 
                                              boxes_per_small_box: int, total_boxes: int = None) -> str:
        """
        生成常规小箱标的序列号范围 - 修复边界计算问题
        对应原来 _create_regular_small_box_label 中的序列号范围计算逻辑
        添加total_boxes边界检查，确保序列号不超出实际盒数
        """
        serial_info = self.parse_serial_number_format(base_number)
        
        # 计算当前小箱包含的盒子范围
        start_box = (small_box_num - 1) * boxes_per_small_box + 1
        end_box = start_box + boxes_per_small_box - 1
        
        # 🔧 边界检查：确保end_box不超过总盒数
        if total_boxes is not None:
            end_box = min(end_box, total_boxes)
        
        # 生成范围内第一个和最后一个序列号
        first_serial_num = serial_info['start_number'] + (start_box - 1)
        last_serial_num = serial_info['start_number'] + (end_box - 1)
        
        first_serial = f"{serial_info['prefix']}{first_serial_num:0{serial_info['digits']}d}"
        last_serial = f"{serial_info['prefix']}{last_serial_num:0{serial_info['digits']}d}"
        
        # 如果首尾序列号相同，只显示一个
        if first_serial == last_serial:
            serial_range = first_serial
        else:
            serial_range = f"{first_serial}-{last_serial}"
        
        print(f"📝 常规小箱标 #{small_box_num}: 包含盒{start_box}-{end_box}, 序列号范围={serial_range}")
        return serial_range
    
    def generate_regular_large_box_serial_range(self, base_number: str, large_box_num: int,
                                              small_boxes_per_large_box: int, boxes_per_small_box: int, total_boxes: int = None) -> str:
        """
        生成常规大箱标的序列号范围 - 修复边界计算问题
        对应原来 _create_regular_large_box_label 中的序列号范围计算逻辑
        添加total_boxes边界检查，确保序列号不超出实际盒数
        """
        serial_info = self.parse_serial_number_format(base_number)
        
        # 计算当前大箱包含的小箱范围
        start_small_box = (large_box_num - 1) * small_boxes_per_large_box + 1
        end_small_box = start_small_box + small_boxes_per_large_box - 1
        
        # 计算当前大箱包含的总盒子范围
        start_box = (start_small_box - 1) * boxes_per_small_box + 1
        end_box = end_small_box * boxes_per_small_box
        
        # 🔧 边界检查：确保end_box不超过总盒数
        if total_boxes is not None:
            end_box = min(end_box, total_boxes)
        
        # 生成范围内第一个和最后一个序列号
        first_serial_num = serial_info['start_number'] + (start_box - 1)
        last_serial_num = serial_info['start_number'] + (end_box - 1)
        
        first_serial = f"{serial_info['prefix']}{first_serial_num:0{serial_info['digits']}d}"
        last_serial = f"{serial_info['prefix']}{last_serial_num:0{serial_info['digits']}d}"
        
        # 如果首尾序列号相同，只显示一个
        if first_serial == last_serial:
            serial_range = first_serial
        else:
            serial_range = f"{first_serial}-{last_serial}"
        
        print(f"📝 常规大箱标 #{large_box_num}: 包含小箱{start_small_box}-{end_small_box}, 盒{start_box}-{end_box}, 序列号范围={serial_range}")
        return serial_range
    
    def calculate_carton_number_for_small_box(self, small_box_num: int, total_small_boxes: int) -> str:
        """计算常规小箱标的Carton No - 格式：第几小箱/总小箱数"""
        return f"{small_box_num}/{total_small_boxes}"
    
    def calculate_carton_range_for_large_box(self, large_box_num: int, total_large_boxes: int) -> str:
        """计算常规大箱标的Carton No - 格式：第几大箱/总大箱数"""
        return f"{large_box_num}/{total_large_boxes}"
    
    def calculate_pieces_for_small_box(self, small_box_num: int, total_small_boxes: int, 
                                     pieces_per_small_box: int, remaining_pieces: int) -> int:
        """
        计算常规小箱的实际数量 - 与原有逻辑完全一致
        处理最后一个小箱可能的余数情况
        """
        if small_box_num == total_small_boxes and remaining_pieces > 0:
            # 最后一个小箱，如果有余数就用余数
            return remaining_pieces
        else:
            # 其他小箱使用标准数量
            return pieces_per_small_box
    
    def calculate_pieces_for_large_box(self, large_box_num: int, total_large_boxes: int,
                                     pieces_per_large_box: int, remaining_pieces: int) -> int:
        """
        计算常规大箱的实际数量 - 与原有逻辑完全一致
        处理最后一个大箱可能的余数情况
        """
        if large_box_num == total_large_boxes and remaining_pieces > 0:
            # 最后一个大箱，如果有余数就用余数
            return remaining_pieces
        else:
            # 其他大箱使用标准数量
            return pieces_per_large_box


# 创建全局实例供regular模板使用
regular_data_processor = RegularDataProcessor()