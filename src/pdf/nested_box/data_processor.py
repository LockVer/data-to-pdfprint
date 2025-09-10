"""
套盒模板专属数据处理器
封装套盒模板的所有数据提取和处理逻辑，确保与原有功能完全一致
"""

import pandas as pd
import re
import math
from typing import Dict, Any

# 导入现有的通用Excel工具，确保功能一致性
from src.utils.excel_data_extractor import ExcelDataExtractor


class NestedBoxDataProcessor:
    """套盒模板专属数据处理器 - 封装现有逻辑，确保功能完全一致"""
    
    def __init__(self):
        """初始化套盒数据处理器"""
        self.template_type = "nested_box"
    
    def extract_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取套盒盒标所需的数据 - 使用统一的公共数据提取方法
        """
        # 使用统一的公共数据提取方法
        extractor = ExcelDataExtractor(excel_file_path)
        return extractor.extract_common_data()
    
    def extract_small_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取套盒套标所需的数据 - 使用统一的公共数据提取方法
        """
        # 使用统一的公共数据提取方法
        extractor = ExcelDataExtractor(excel_file_path)
        return extractor.extract_common_data()
    
    def extract_large_box_label_data(self, excel_file_path: str) -> Dict[str, Any]:
        """
        提取套盒箱标所需的数据 - 使用统一的公共数据提取方法
        """
        # 使用统一的公共数据提取方法
        extractor = ExcelDataExtractor(excel_file_path)
        return extractor.extract_common_data()
    
    def _fallback_keyword_extraction_for_box_label(self, excel_file_path: str) -> Dict[str, Any]:
        """回退到关键字提取方式 - 与原代码逻辑完全一致"""
        extractor = ExcelDataExtractor(excel_file_path)
        keyword_config = {
            '标签名称': {'keyword': '标签名称', 'direction': 'right'},
            '开始号': {'keyword': '开始号', 'direction': 'down'},
            '结束号': {'keyword': '结束号', 'direction': 'down'},
            '客户编码': {'keyword': '客户名称编码', 'direction': 'down'}
        }
        excel_data = extractor.extract_data_by_keywords(keyword_config)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        end_number = excel_data.get('结束号') or base_number
        
        return {
            '标签名称': theme_text,
            '开始号': base_number,
            '结束号': end_number,
            '主题': 'Unknown Theme'
        }
    
    def parse_serial_number_format(self, serial_number: str) -> Dict[str, Any]:
        """
        解析序列号格式 - 与原有逻辑完全一致
        对应原来代码中的正则表达式解析逻辑
        """
        match = re.search(r'(.+?)(\d+)-(\d+)', serial_number)
        
        if match:
            start_prefix = match.group(1)
            start_main = int(match.group(2))
            start_suffix = int(match.group(3))
            
            print(f"✅ 解析序列号格式:")
            print(f"   开始: {start_prefix}{start_main:05d}-{start_suffix:02d}")
            
            return {
                'prefix': start_prefix,
                'main_number': start_main,
                'suffix': start_suffix,
                'main_digits': 5,  # 与原代码一致
                'suffix_digits': 2  # 与原代码一致
            }
        else:
            print("⚠️ 无法解析序列号格式，使用默认逻辑")
            return {
                'prefix': "JAW",
                'main_number': 1001,
                'suffix': 1,
                'main_digits': 5,
                'suffix_digits': 2
            }
    
    def calculate_quantities(self, total_pieces: int, pieces_per_box: int, 
                           boxes_per_small_box: int, small_boxes_per_large_box: int) -> Dict[str, int]:
        """
        计算各级数量 - 与原有逻辑完全一致
        对应原来各个方法中的数量计算逻辑
        """
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
    
    def generate_box_serial_number(self, base_number: str, box_num: int, boxes_per_ending_unit: int) -> str:
        """
        生成套盒盒标的序列号 - 与原有逻辑完全一致
        对应原来 _create_nested_box_label 中的序列号生成逻辑
        """
        serial_info = self.parse_serial_number_format(base_number)
        
        # 套盒模板序列号生成逻辑 - 基于开始号和结束号范围（与原代码完全一致）
        box_index = box_num - 1
        
        # 计算当前盒的序列号在范围内的位置
        main_offset = box_index // boxes_per_ending_unit
        suffix_in_range = (box_index % boxes_per_ending_unit) + serial_info['suffix']
        
        current_main = serial_info['main_number'] + main_offset
        current_number = f"{serial_info['prefix']}{current_main:05d}-{suffix_in_range:02d}"
        
        print(f"📝 生成套盒盒标 #{box_num}: {current_number}")
        return current_number
    
    def generate_small_box_serial_range(self, base_number: str, small_box_num: int, boxes_per_small_box: int, total_boxes: int = None) -> str:
        """
        生成套盒套标的序列号范围 - 修复边界计算问题
        对应原来 _create_nested_small_box_label 中的序列号范围计算逻辑
        添加total_boxes边界检查，确保序列号不超出实际盒数
        """
        match = re.search(r'(\d+)', base_number)
        if match:
            # 获取第一个数字（主号）的起始位置
            digit_start = match.start()
            # 截取主号前面的所有字符作为前缀
            prefix_part = base_number[:digit_start]
            base_main_num = int(match.group(1))  # 主号
            
            # 套盒模板套标的简化逻辑：
            # 每个套标对应一个主号，包含连续的boxes_per_small_box个副号
            current_main_number = base_main_num + (small_box_num - 1)  # 当前套对应的主号
            
            # 计算当前套实际包含的盒数范围
            start_box = (small_box_num - 1) * boxes_per_small_box + 1
            end_box = start_box + boxes_per_small_box - 1
            
            # 🔧 边界检查：确保end_box不超过总盒数
            if total_boxes is not None:
                end_box = min(end_box, total_boxes)
            
            # 副号从01开始，根据实际盒数计算结束副号
            start_suffix = 1
            actual_boxes_in_small_box = end_box - start_box + 1
            end_suffix = start_suffix + actual_boxes_in_small_box - 1
            
            start_serial = f"{prefix_part}{current_main_number:05d}-{start_suffix:02d}"
            end_serial = f"{prefix_part}{current_main_number:05d}-{end_suffix:02d}"
            
            # 套盒套标显示序列号范围
            if start_suffix == end_suffix:
                serial_range = start_serial
            else:
                serial_range = f"{start_serial}-{end_serial}"
                
            print(f"📝 套盒套标 #{small_box_num}: 主号{current_main_number}, 副号{start_suffix}-{end_suffix}, 包含盒{start_box}-{end_box} = {serial_range}")
            return serial_range
        else:
            return f"DSK{small_box_num:05d}-DSK{small_box_num:05d}"
    
    def generate_large_box_serial_range(self, base_number: str, large_box_num: int, 
                                      small_boxes_per_large_box: int, boxes_per_small_box: int, total_boxes: int = None) -> str:
        """
        生成套盒箱标的序列号范围 - 修复边界计算问题
        对应原来 _create_nested_large_box_label 中的序列号范围计算逻辑
        添加total_boxes边界检查，确保序列号不超出实际盒数
        """
        # 计算当前箱包含的套范围
        start_small_box = (large_box_num - 1) * small_boxes_per_large_box + 1
        end_small_box = start_small_box + small_boxes_per_large_box - 1
        
        # 计算当前箱包含的总盒子范围
        start_box = (start_small_box - 1) * boxes_per_small_box + 1
        end_box = end_small_box * boxes_per_small_box
        
        # 🔧 边界检查：确保end_box不超过总盒数
        if total_boxes is not None:
            end_box = min(end_box, total_boxes)
            # 根据实际的end_box重新计算最后一个套
            actual_end_small_box = math.ceil(end_box / boxes_per_small_box)
            end_small_box = min(end_small_box, actual_end_small_box)
        
        # 计算序列号范围 - 从第一个套的起始号到最后一个套的结束号
        match = re.search(r'(\d+)', base_number)
        if match:
            # 获取第一个数字（主号）的起始位置
            digit_start = match.start()
            # 截取主号前面的所有字符作为前缀
            prefix_part = base_number[:digit_start]
            base_main_num = int(match.group(1))  # 主号
            
            # 第一个套的序列号范围
            first_main_number = base_main_num + (start_small_box - 1)
            first_start_serial = f"{prefix_part}{first_main_number:05d}-01"
            
            # 最后一个套的序列号范围（考虑边界）
            last_main_number = base_main_num + (end_small_box - 1)
            last_box_in_small_box = end_box - (end_small_box - 1) * boxes_per_small_box
            last_suffix = min(boxes_per_small_box, last_box_in_small_box)
            last_end_serial = f"{prefix_part}{last_main_number:05d}-{last_suffix:02d}"
            
            # 箱标显示完整序列号范围
            if first_start_serial == last_end_serial:
                serial_range = first_start_serial
            else:
                serial_range = f"{first_start_serial}-{last_end_serial}"
                
            print(f"📝 套盒箱标 #{large_box_num}: 包含套{start_small_box}-{end_small_box}, 盒{start_box}-{end_box}, 序列号范围={serial_range}")
            return serial_range
        else:
            return f"DSK{large_box_num:05d}-DSK{large_box_num:05d}"
    
    def calculate_carton_number_for_small_box(self, small_box_num: int) -> str:
        """计算套盒套标的Carton No - 与原有逻辑完全一致"""
        return str(small_box_num)
    
    def calculate_carton_range_for_large_box(self, large_box_num: int, small_boxes_per_large_box: int) -> str:
        """计算套盒箱标的Carton No范围 - 与原有逻辑完全一致"""
        start_small_box = (large_box_num - 1) * small_boxes_per_large_box + 1
        end_small_box = start_small_box + small_boxes_per_large_box - 1
        return f"{start_small_box}-{end_small_box}"


# 创建全局实例供nested_box模板使用
nested_box_data_processor = NestedBoxDataProcessor()