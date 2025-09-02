"""
Excel文件读取器

负责读取Excel文件并解析数据
"""

import pandas as pd
import os
from typing import Dict, Any, Optional
from pathlib import Path

class ExcelReader:
    """
    Excel文件读取器类
    """
    
    def __init__(self, file_path: str):
        """
        初始化Excel读取器
        
        Args:
            file_path: Excel文件路径
        """
        self.file_path = Path(file_path)
        if not self.file_path.exists():
            raise FileNotFoundError(f"Excel file not found: {file_path}")
    
    def read_data(self) -> pd.DataFrame:
        """
        读取Excel数据
        
        Returns:
            解析后的数据
        """
        try:
            return pd.read_excel(self.file_path)
        except Exception as e:
            raise Exception(f"Error reading Excel file {self.file_path}: {e}")
    
    def extract_template_variables(self) -> Dict[str, Any]:
        """
        从Excel模板中提取关键变量
        
        Returns:
            包含客户编码、主题、开始号、总张数等的字典
        """
        df = self.read_data()
        variables = {}
        
        # 根据实际Excel结构读取数据
        try:
            # A3: 客户名称编码 (实际位置在第3行第0列，索引为2,0)
            if len(df) > 2 and len(df.columns) > 0:
                variables['customer_code'] = str(df.iloc[2, 0]) if pd.notna(df.iloc[2, 0]) else ''
            
            # B3: 主题 (实际位置在第3行第1列，索引为2,1) - 只保留英文部分  
            if len(df) > 2 and len(df.columns) > 1:
                theme_full = str(df.iloc[2, 1]) if pd.notna(df.iloc[2, 1]) else ''
                # 简单提取英文部分：删除中文字符、多余空格和连字符
                import re
                # 删除中文字符（Unicode范围）
                english_only = re.sub(r'[\u4e00-\u9fff]', '', theme_full)
                # 删除连字符
                english_only = english_only.replace('-', '')
                # 清理多余空格
                english_only = re.sub(r'\s+', ' ', english_only).strip()
                
                variables['theme'] = english_only if english_only else theme_full
            
            # 查找开始号：在B9或B10位置查找
            variables['start_number'] = '1'  # 默认值
            if len(df) > 8 and len(df.columns) > 1:
                # 先检查B10 (索引9,1)
                if pd.notna(df.iloc[9, 1]):
                    start_val = str(df.iloc[9, 1])
                    # 检查是否包含字母（如LGM01001）
                    import re
                    if re.search(r'[A-Za-z]', start_val):
                        variables['start_number'] = start_val
                # 如果B10没有，检查B9 (索引8,1)  
                elif pd.notna(df.iloc[8, 1]):
                    start_val = str(df.iloc[8, 1])
                    if re.search(r'[A-Za-z]', start_val):
                        variables['start_number'] = start_val
            
            # F3: 总张数 (实际位置在第3行第5列，索引为2,5)
            if len(df) > 2 and len(df.columns) > 5:
                total_sheets_val = df.iloc[2, 5]
                if pd.notna(total_sheets_val):
                    # 提取数字部分
                    import re
                    numbers = re.findall(r'\d+', str(total_sheets_val))
                    variables['total_sheets'] = int(numbers[0]) if numbers else 1
                else:
                    variables['total_sheets'] = 1
                    
        except (IndexError, ValueError) as e:
            # 如果读取失败，设置默认值
            variables.setdefault('customer_code', '')
            variables.setdefault('theme', '')
            variables.setdefault('start_number', '1')
            variables.setdefault('total_sheets', 1)
        
        return variables
    
    def detect_packaging_mode(self) -> str:
        """
        根据文件名检测包装模式
        
        Returns:
            packaging mode: 'regular', 'separate_box', 'set_box'
        """
        filename = self.file_path.name.lower()
        
        if '常规' in filename:
            return 'regular'
        elif '分盒' in filename:
            return 'separate_box'
        elif '套盒' in filename:
            return 'set_box'
        else:
            # 默认为常规模式
            return 'regular'


class PackagingModeDetector:
    """
    包装模式检测器
    """
    
    @staticmethod
    def detect_from_path(file_path: str) -> str:
        """
        从文件路径检测包装模式
        
        Args:
            file_path: Excel文件路径
            
        Returns:
            包装模式: 'regular', 'separate_box', 'set_box'
        """
        path = Path(file_path)
        
        # 检查路径中的目录名
        if '常规' in str(path):
            return 'regular'
        elif '分盒' in str(path):
            return 'separate_box'
        elif '套盒' in str(path):
            return 'set_box'
        
        # 检查文件名
        filename = path.name.lower()
        if '常规' in filename:
            return 'regular'
        elif '分盒' in filename:
            return 'separate_box'
        elif '套盒' in filename:
            return 'set_box'
        
        return 'regular'  # 默认模式