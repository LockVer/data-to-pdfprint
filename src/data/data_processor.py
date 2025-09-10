"""
数据处理器

负责处理和转换Excel中的原始数据
"""

import math
from typing import Dict, Any, List
from dataclasses import dataclass


@dataclass
class PackagingConfig:
    """包装配置参数"""
    box_quantity: int  # 分盒张数
    set_quantity: int  # 分套张数 (仅套盒模式使用) - 将被弃用，改用计算值
    small_box_capacity: int  # 小箱内的盒数
    large_box_capacity: int  # 大箱内的小箱数
    cards_per_box_in_set: int  # 套中每盒张数 (仅套盒模式使用)
    boxes_per_set: int  # 每套盒数 (仅套盒模式使用)
    is_overweight: bool = False  # 是否超重模式 (仅套盒模式使用)
    sets_per_large_box: int = 2  # 每大箱套数 (仅套盒非超重模式使用)
    cases_per_set: int = 1  # 一套分为几箱 (仅套盒超重模式使用)


class DataProcessor:
    """
    数据处理器类
    """
    
    def __init__(self):
        """
        初始化数据处理器
        """
        pass
    
    def extract_fields(self, raw_data):
        """
        提取指定字段
        
        Args:
            raw_data: 原始数据
            
        Returns:
            提取后的字段数据
        """
        pass
    
    def validate_data(self, data):
        """
        验证数据格式
        
        Args:
            data: 待验证的数据
            
        Returns:
            验证结果
        """
        pass
    
    def process_for_packaging_mode(self, variables: Dict[str, Any], config: PackagingConfig, 
                                   packaging_mode: str) -> Dict[str, Any]:
        """
        根据包装模式处理数据
        
        Args:
            variables: 从Excel提取的变量
            config: 包装配置参数
            packaging_mode: 包装模式
            
        Returns:
            处理后的数据，包含计算的标签信息
        """
        base_data = {
            'customer_code': variables.get('customer_code', ''),
            'theme': variables.get('theme', ''),
            'start_number': variables.get('start_number', ''),
            'packaging_mode': packaging_mode
        }
        
        if packaging_mode == 'regular':
            return self._process_regular_mode(base_data, config)
        elif packaging_mode == 'separate_box':
            return self._process_separate_box_mode(base_data, config)
        elif packaging_mode == 'set_box':
            return self._process_set_box_mode(base_data, config)
        else:
            raise ValueError(f"Unknown packaging mode: {packaging_mode}")
    
    def _process_regular_mode(self, base_data: Dict[str, Any], config: PackagingConfig) -> Dict[str, Any]:
        """
        处理常规模式
        常规箱标（一盒入箱再两盒入一箱）
        """
        box_quantity = config.box_quantity
        
        # 计算箱数：一盒入一小箱，两小箱入一大箱
        small_boxes = box_quantity  # 每盒一个小箱
        large_boxes = math.ceil(small_boxes / 2)  # 两小箱入一大箱
        
        result = {
            **base_data,
            'box_quantity': box_quantity,
            'small_box_quantity': small_boxes,
            'large_box_quantity': large_boxes,
            'boxes_per_small_box': 1,
            'small_boxes_per_large_box': 2,
            'label_specifications': {
                'box_labels': box_quantity,
                'small_box_labels': small_boxes,
                'large_box_labels': large_boxes
            }
        }
        
        return result
    
    def _process_separate_box_mode(self, base_data: Dict[str, Any], config: PackagingConfig) -> Dict[str, Any]:
        """
        处理分盒模式
        分盒箱标（根据用户输入的每小箱盒数进行计算）
        """
        box_quantity = config.box_quantity
        boxes_per_small_box = config.small_box_capacity  # 使用用户输入的每小箱盒数
        small_boxes_per_large_box = config.large_box_capacity  # 使用用户输入的每大箱小箱数
        
        # 分盒模式：根据用户输入计算小箱和大箱数量
        small_boxes = math.ceil(box_quantity / boxes_per_small_box) if boxes_per_small_box > 0 else box_quantity
        large_boxes = math.ceil(small_boxes / small_boxes_per_large_box) if small_boxes_per_large_box > 0 else small_boxes
        
        result = {
            **base_data,
            'box_quantity': box_quantity,
            'small_box_quantity': small_boxes,
            'large_box_quantity': large_boxes,
            'boxes_per_small_box': boxes_per_small_box,
            'small_boxes_per_large_box': small_boxes_per_large_box,
            'label_specifications': {
                'box_labels': box_quantity,
                'small_box_labels': small_boxes,
                'large_box_labels': large_boxes
            }
        }
        
        return result
    
    def _process_set_box_mode(self, base_data: Dict[str, Any], config: PackagingConfig) -> Dict[str, Any]:
        """
        处理套盒模式
        套盒箱标（根据用户输入的套中每盒张数和每套盒数计算每套张数，支持超重和非超重模式）
        """
        box_quantity = config.box_quantity
        cards_per_box_in_set = config.cards_per_box_in_set
        boxes_per_set = config.boxes_per_set
        is_overweight = config.is_overweight
        
        # 计算每套张数：套中每盒张数 × 每套盒数
        cards_per_set = cards_per_box_in_set * boxes_per_set
        
        # 计算总套数
        total_sets = math.ceil(box_quantity / boxes_per_set)
        
        if is_overweight:
            # 超重模式：一套分为几箱
            cases_per_set = config.cases_per_set
            total_cases = total_sets * cases_per_set  # 总箱数：每套分为多个箱
            
            # 计算每箱中盒数
            boxes_per_case = math.ceil(boxes_per_set / cases_per_set)
            
            result = {
                **base_data,
                'box_quantity': box_quantity,
                'cards_per_box_in_set': cards_per_box_in_set,
                'boxes_per_set': boxes_per_set,
                'cards_per_set': cards_per_set,
                'total_sets': total_sets,
                'total_cases': total_cases,
                'is_overweight': is_overweight,
                'cases_per_set': cases_per_set,
                'boxes_per_case': boxes_per_case,
                'label_specifications': {
                    'box_labels': box_quantity,
                    'case_labels': total_cases  # 超重模式只有箱标
                }
            }
        else:
            # 非超重模式：每大箱套数
            sets_per_large_box = config.sets_per_large_box
            small_boxes = total_sets  # 每套一个小箱
            large_boxes = math.ceil(total_sets / sets_per_large_box)  # 根据每大箱套数计算大箱数
            sets_per_small_box = 1
            
            result = {
                **base_data,
                'box_quantity': box_quantity,
                'cards_per_box_in_set': cards_per_box_in_set,
                'boxes_per_set': boxes_per_set,
                'cards_per_set': cards_per_set,
                'total_sets': total_sets,
                'small_box_quantity': small_boxes,
                'large_box_quantity': large_boxes,
                'is_overweight': is_overweight,
                'sets_per_large_box': sets_per_large_box,
                'sets_per_small_box': sets_per_small_box,
                'label_specifications': {
                    'box_labels': box_quantity,
                    'set_labels': total_sets,
                    'small_box_labels': small_boxes,
                    'large_box_labels': large_boxes
                }
            }
        
        return result
    
    def calculate_label_ranges(self, processed_data: Dict[str, Any]) -> Dict[str, List[str]]:
        """
        计算标签编号范围
        
        Args:
            processed_data: 处理后的数据
            
        Returns:
            各类型标签的编号列表
        """
        start_number = processed_data.get('start_number', '')
        packaging_mode = processed_data.get('packaging_mode', 'regular')
        
        # 提取起始编号的数字部分
        import re
        match = re.search(r'(\d+)', start_number)
        if not match:
            start_num = 1
        else:
            start_num = int(match.group(1))
        
        # 提取编号前缀
        prefix = start_number.replace(match.group(1), '') if match else ''
        
        ranges = {}
        
        if packaging_mode == 'set_box':
            # 套盒模式的编号生成
            box_quantity = processed_data.get('box_quantity', 0)
            set_quantity = processed_data.get('set_quantity', 6)
            
            ranges['box_numbers'] = [
                f"{prefix}{start_num + i:05d}" for i in range(box_quantity)
            ]
            
            # 套装编号
            total_sets = processed_data.get('total_sets', 0)
            ranges['set_numbers'] = [
                f"{prefix}SET{i+1:02d}" for i in range(total_sets)
            ]
            
        else:
            # 常规和分盒模式
            box_quantity = processed_data.get('box_quantity', 0)
            ranges['box_numbers'] = [
                f"{prefix}{start_num + i:05d}" for i in range(box_quantity)
            ]
        
        return ranges