"""
通用序列号格式化工具
提供智能序列号解析和格式化功能，保持原始数字位数
"""

import re
from typing import Dict, Any


class SerialNumberFormatter:
    """通用序列号格式化工具类"""
    
    @staticmethod
    def parse_serial_number_format(serial_number: str) -> Dict[str, Any]:
        """
        解析序列号格式 - 保持原始数字格式
        
        参数:
            serial_number: 输入的序列号字符串 (如 MCH0102, DSK01001)
            
        返回:
            包含前缀、数字值、原始位数等信息的字典
        """
        if not serial_number:
            return {
                'prefix': 'DSK',
                'main_number': 1001,
                'original_digits': 5,
                'digit_start': 0
            }
        
        # 查找第一个数字序列
        match = re.search(r'(\d+)', serial_number)
        if match:
            # 获取第一个数字（主号）的起始位置
            digit_start = match.start()
            # 截取主号前面的所有字符作为前缀
            prefix_part = serial_number[:digit_start]
            original_number_str = match.group(1)  # 原始数字字符串
            base_main_num = int(original_number_str)  # 主号数值
            
            return {
                'prefix': prefix_part,
                'main_number': base_main_num,
                'original_digits': len(original_number_str),  # 🔑 保持原始位数
                'digit_start': digit_start
            }
        else:
            return {
                'prefix': 'DSK',
                'main_number': 1001,
                'original_digits': 5,  # 默认5位
                'digit_start': 0
            }
    
    @staticmethod
    def format_serial_number(prefix: str, number: int, original_digits: int) -> str:
        """
        智能格式化序列号 - 保持原始位数
        
        参数:
            prefix: 前缀字符串
            number: 数字值
            original_digits: 原始数字位数
            
        返回:
            格式化后的序列号，保持原始位数
            
        示例:
            format_serial_number("MCH", 102, 4) -> "MCH0102"
            format_serial_number("DSK", 1001, 5) -> "DSK01001"
            format_serial_number("BOX", 123, 3) -> "BOX123"
        """
        # 根据原始位数决定格式化方式
        return f"{prefix}{number:0{original_digits}d}"
    
    @staticmethod
    def migrate_legacy_formatting(legacy_format_string: str, prefix: str, number: int, original_digits: int) -> str:
        """
        迁移旧版格式化代码的辅助函数
        将 f"{prefix}{number:05d}" 这样的调用替换为智能格式化
        
        参数:
            legacy_format_string: 原始的格式化字符串（用于调试）
            prefix: 前缀字符串
            number: 数字值
            original_digits: 原始数字位数
            
        返回:
            使用新格式化逻辑的结果
        """
        return SerialNumberFormatter.format_serial_number(prefix, number, original_digits)


# 创建全局实例供其他模块使用
serial_formatter = SerialNumberFormatter()