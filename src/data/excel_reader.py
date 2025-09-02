"""
Excel文件读取器

负责读取Excel文件并解析数据
"""

import pandas as pd
import os
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
        self.data = None
        
    def read_data(self):
        """
        读取Excel数据
        
        Returns:
            解析后的数据
        """
        try:
            if not self.file_path.exists():
                raise FileNotFoundError(f"文件不存在: {self.file_path}")
            
            # 读取Excel文件
            if self.file_path.suffix.lower() == '.xlsx':
                self.data = pd.read_excel(self.file_path, engine='openpyxl')
            elif self.file_path.suffix.lower() == '.xls':
                self.data = pd.read_excel(self.file_path, engine='xlrd')
            else:
                raise ValueError("不支持的文件格式，请使用 .xlsx 或 .xls 文件")
            
            return self.data
        except Exception as e:
            raise Exception(f"读取Excel文件失败: {str(e)}")
    
    def get_file_info(self):
        """
        获取文件信息
        
        Returns:
            文件信息字典
        """
        if not self.file_path.exists():
            return None
        
        stat = self.file_path.stat()
        return {
            'name': self.file_path.name,
            'size': stat.st_size,
            'size_mb': round(stat.st_size / (1024*1024), 2),
            'rows': len(self.data) if self.data is not None else 0
        }
    
    def get_columns(self):
        """
        获取Excel文件的列名
        
        Returns:
            列名列表
        """
        if self.data is not None:
            return list(self.data.columns)
        return []