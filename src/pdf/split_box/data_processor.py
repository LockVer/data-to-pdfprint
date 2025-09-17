"""
分盒模板专属数据处理器
封装分盒模板的所有数据提取和处理逻辑，确保与原有功能完全一致
"""

import pandas as pd
import re
import math
from typing import Dict, Any

# 导入现有的通用Excel工具，确保功能一致性
from src.utils.excel_data_extractor import ExcelDataExtractor


class SplitBoxDataProcessor:
    """分盒模板专属数据处理器 - 封装现有逻辑，确保功能完全一致"""
    
    def __init__(self):
        """初始化分盒数据处理器"""
        self.template_type = "split_box"
    
    def extract_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取分盒盒标所需的数据 - 使用统一的公共数据提取方法
        """
        # 使用统一的公共数据提取方法
        extractor = ExcelDataExtractor(excel_file_path)
        return extractor.extract_common_data()
    
    def extract_small_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取分盒小箱标所需的数据 - 使用统一的公共数据提取方法
        """
        # 使用统一的公共数据提取方法
        extractor = ExcelDataExtractor(excel_file_path)
        return extractor.extract_common_data()
    
    def extract_large_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取分盒大箱标所需的数据 - 使用统一的公共数据提取方法
        """
        # 使用统一的公共数据提取方法
        extractor = ExcelDataExtractor(excel_file_path)
        return extractor.extract_common_data()
    
    def parse_serial_number_format(self, serial_number: str) -> Dict[str, Any]:
        """
        解析序列号格式 - 与原有逻辑完全一致
        对应原来代码中的正则表达式解析逻辑
        """
        # 分盒模板使用简单的数字提取逻辑（与原代码一致）
        match = re.search(r'(\d+)', serial_number)
        if match:
            # 获取第一个数字（主号）的起始位置
            digit_start = match.start()
            # 截取主号前面的所有字符作为前缀
            prefix_part = serial_number[:digit_start]
            base_main_num = int(match.group(1))  # 主号
            
            return {
                'prefix': prefix_part,
                'main_number': base_main_num,
                'digit_start': digit_start
            }
        else:
            return {
                'prefix': 'DSK',
                'main_number': 1001,
                'digit_start': 0
            }
    
    def calculate_quantities(self, total_pieces: int, pieces_per_box: int, 
                           boxes_per_small_box: int, small_boxes_per_large_box: int) -> Dict[str, int]:
        """
        计算各级数量 - 与原有逻辑完全一致
        对应原来各个方法中的数量计算逻辑  
        """
        # 计算数量 - 三级结构：张→盒→小箱→大箱（使用向上取整处理余数）
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)
        total_large_boxes = math.ceil(total_small_boxes / small_boxes_per_large_box)
        
        return {
            'total_pieces': total_pieces,
            'total_boxes': total_boxes,
            'total_small_boxes': total_small_boxes,
            'total_large_boxes': total_large_boxes,
            'pieces_per_box': pieces_per_box,
            'boxes_per_small_box': boxes_per_small_box,
            'small_boxes_per_large_box': small_boxes_per_large_box
        }
    
    def generate_split_box_serial_number(self, base_number: str, box_num: int, boxes_per_small_box: int, small_boxes_per_large_box: int) -> str:
        """
        生成分盒盒标的序列号 - 修正为文档中的分盒逻辑
        对应原来 _create_split_box_label 中的序列号生成逻辑
        
        分盒模板特殊逻辑：副号进位阈值 = 盒/小箱 × 小箱/大箱
        """
        # 计算副号进位阈值
        group_size = boxes_per_small_box * small_boxes_per_large_box
        serial_info = self.parse_serial_number_format(base_number)
        
        # 分盒模板序列号生成逻辑（与原代码完全一致）
        box_index = box_num - 1  # 转换为0-based索引
        
        # 分盒模板的特殊逻辑：副号满group_size进一
        main_increments = box_index // group_size  # 主号增加的次数
        suffix_in_group = (box_index % group_size) + 1  # 当前组内的副号（1-based）
        
        current_main = serial_info['main_number'] + main_increments
        current_number = f"{serial_info['prefix']}{current_main:05d}-{suffix_in_group:02d}"
        
        print(f"📝 分盒盒标 #{box_num}: 主号{current_main}, 副号{suffix_in_group}, 分组大小{group_size}({boxes_per_small_box}×{small_boxes_per_large_box}) → {current_number}")
        return current_number
    
    def generate_split_small_box_serial_range(self, base_number: str, small_box_num: int, 
                                            boxes_per_small_box: int, small_boxes_per_large_box: int, total_boxes: int = None) -> str:
        """
        生成分盒小箱标的序列号范围 - 修复边界计算问题，使用正确的副号进位阈值
        对应原来 _create_split_small_box_label 中的序列号范围计算逻辑
        添加total_boxes边界检查，确保序列号不超出实际盒数
        """
        # 计算副号进位阈值
        group_size = boxes_per_small_box * small_boxes_per_large_box
        serial_info = self.parse_serial_number_format(base_number)
        
        # 计算当前小箱包含的盒子范围
        start_box = (small_box_num - 1) * boxes_per_small_box + 1
        end_box = start_box + boxes_per_small_box - 1
        
        # 🔧 边界检查：确保end_box不超过总盒数
        if total_boxes is not None:
            end_box = min(end_box, total_boxes)
        
        # 计算范围内第一个盒子的序列号
        first_box_index = start_box - 1
        first_main_increments = first_box_index // group_size
        first_suffix = (first_box_index % group_size) + 1
        first_main = serial_info['main_number'] + first_main_increments
        first_serial = f"{serial_info['prefix']}{first_main:05d}-{first_suffix:02d}"
        
        # 计算范围内最后一个盒子的序列号  
        last_box_index = end_box - 1
        last_main_increments = last_box_index // group_size
        last_suffix = (last_box_index % group_size) + 1
        last_main = serial_info['main_number'] + last_main_increments
        last_serial = f"{serial_info['prefix']}{last_main:05d}-{last_suffix:02d}"
        
        # 始终显示为范围格式，即使首尾序列号相同
        serial_range = f"{first_serial}-{last_serial}"
        
        print(f"📝 分盒小箱标 #{small_box_num}: 包含盒{start_box}-{end_box}, 序列号范围={serial_range}")
        return serial_range
    
    def generate_split_large_box_serial_range(self, base_number: str, large_box_num: int,
                                            small_boxes_per_large_box: int, boxes_per_small_box: int, 
                                            total_boxes: int = None) -> str:
        """
        生成分盒大箱标的序列号范围 - 修复边界计算问题，使用正确的副号进位阈值
        对应原来 _create_split_large_box_label 中的序列号范围计算逻辑
        添加total_boxes边界检查，确保序列号不超出实际盒数
        """
        # 计算副号进位阈值
        group_size = boxes_per_small_box * small_boxes_per_large_box
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
        
        # 计算范围内第一个盒子的序列号
        first_box_index = start_box - 1
        first_main_increments = first_box_index // group_size
        first_suffix = (first_box_index % group_size) + 1
        first_main = serial_info['main_number'] + first_main_increments
        first_serial = f"{serial_info['prefix']}{first_main:05d}-{first_suffix:02d}"
        
        # 计算范围内最后一个盒子的序列号
        last_box_index = end_box - 1
        last_main_increments = last_box_index // group_size
        last_suffix = (last_box_index % group_size) + 1
        last_main = serial_info['main_number'] + last_main_increments
        last_serial = f"{serial_info['prefix']}{last_main:05d}-{last_suffix:02d}"
        
        # 始终显示为范围格式，即使首尾序列号相同
        serial_range = f"{first_serial}-{last_serial}"
        
        print(f"📝 分盒大箱标 #{large_box_num}: 包含小箱{start_small_box}-{end_small_box}, 盒{start_box}-{end_box}, 序列号范围={serial_range}")
        return serial_range
    
    def calculate_carton_number_for_small_box(self, small_box_num: int, boxes_per_set: int, boxes_per_small_box: int) -> str:
        """
        计算分盒小箱标的Carton No - 基于最新逻辑整理
        根据每套小箱数量判断渲染模式
        """
        # 计算每套小箱数量
        small_boxes_per_set = boxes_per_set / boxes_per_small_box
        
        if small_boxes_per_set > 1:
            # 一套分多个小箱：多级编号 (套号-小箱号)
            set_num = ((small_box_num - 1) // int(small_boxes_per_set)) + 1
            small_box_in_set = ((small_box_num - 1) % int(small_boxes_per_set)) + 1
            return f"{set_num}-{small_box_in_set}"
        elif small_boxes_per_set == 1:
            # 一套分一个小箱：单级编号 (01, 02, 03...)
            return f"{small_box_num:02d}"
        else:
            # 没有小箱标 (每套小箱数量 = null/0)
            return None
    
    def calculate_carton_range_for_large_box(self, large_box_num: int, boxes_per_set: int, boxes_per_large_box: int, total_sets: int) -> str:
        """
        计算分盒大箱标的Carton No - 基于最新逻辑整理
        根据每套大箱数量判断渲染模式
        """
        # 计算每套大箱数量
        large_boxes_per_set = boxes_per_set / boxes_per_large_box
        
        if large_boxes_per_set > 1:
            # 一套分多个大箱：多级编号 (套号-大箱号)
            set_num = ((large_box_num - 1) // int(large_boxes_per_set)) + 1
            large_box_in_set = ((large_box_num - 1) % int(large_boxes_per_set)) + 1
            return f"{set_num}-{large_box_in_set}"
        elif large_boxes_per_set == 1:
            # 一套分一个大箱：单级编号 (1, 2, 3...)
            return str(large_box_num)
        else:
            # 多套分一个大箱：多级编号 (起始套号-结束套号)
            sets_per_large_box = int(1 / large_boxes_per_set)
            start_set = (large_box_num - 1) * sets_per_large_box + 1
            end_set = min(start_set + sets_per_large_box - 1, total_sets)
            return f"{start_set}-{end_set}"
    
    def calculate_group_size(self, boxes_per_small_box: int, small_boxes_per_large_box: int) -> int:
        """
        计算副号进位阈值 - 修正为文档中的逻辑
        副号进位阈值 = 盒/小箱 × 小箱/大箱
        """
        group_size = boxes_per_small_box * small_boxes_per_large_box
        print(f"✅ 分盒模板副号进位阈值: {group_size} (盒/小箱{boxes_per_small_box} × 小箱/大箱{small_boxes_per_large_box})")
        return group_size


# 创建全局实例供split_box模板使用
split_box_data_processor = SplitBoxDataProcessor()