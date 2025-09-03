"""
Excel文件读取器

负责读取Excel文件并解析数据
"""

import pandas as pd
import openpyxl
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
    
    def extract_box_label_data(self, sheet_name=None):
        """
        提取盒标相关的特定单元格数据
        
        Args:
            sheet_name: 工作表名称，如果为None则使用第一个工作表
        
        Returns:
            dict: 包含A4, B4, B11, F4位置数据的字典，并尝试计算结束号
        """
        try:
            if not self.file_path.exists():
                raise FileNotFoundError(f"文件不存在: {self.file_path}")
            
            # 使用openpyxl直接读取特定单元格
            workbook = openpyxl.load_workbook(self.file_path, data_only=True)
            if sheet_name:
                worksheet = workbook[sheet_name]
            else:
                worksheet = workbook.active
            
            # 提取指定单元格数据
            extracted_data = {
                'A4': worksheet['A4'].value,   # 客户名称编码
                'B4': worksheet['B4'].value,   # 主题
                'B11': worksheet['B11'].value, # 开始号
                'F4': worksheet['F4'].value,   # 总张数
            }
            
            # 清理数据 - 移除None值并转换为适当格式
            for key, value in extracted_data.items():
                if value is None:
                    extracted_data[key] = ""
                else:
                    extracted_data[key] = str(value).strip()
            
            # 处理总张数，确保是数字
            try:
                if extracted_data['F4']:
                    extracted_data['F4'] = int(float(extracted_data['F4']))
                else:
                    extracted_data['F4'] = 0
            except (ValueError, TypeError):
                extracted_data['F4'] = 0
            
            # 常规模版1：简单的+1递增，不需要复杂分析
            
            workbook.close()
            return extracted_data
            
        except Exception as e:
            raise Exception(f"提取盒标数据失败: {str(e)}")
    
    def get_worksheet_names(self):
        """
        获取Excel文件中所有工作表名称
        
        Returns:
            list: 工作表名称列表
        """
        try:
            if not self.file_path.exists():
                return []
            
            workbook = openpyxl.load_workbook(self.file_path)
            sheet_names = workbook.sheetnames
            workbook.close()
            return sheet_names
            
        except Exception as e:
            print(f"获取工作表名称失败: {e}")
            return []