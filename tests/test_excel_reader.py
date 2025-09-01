"""
Excel读取器测试
"""

import pytest
import pandas as pd
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
from src.data.excel_reader import ExcelReader


class TestExcelReader:
    """ExcelReader类的测试用例"""
    
    @patch('pathlib.Path.exists')
    def test_init(self, mock_exists):
        """测试初始化"""
        mock_exists.return_value = True
        reader = ExcelReader("test.xlsx")
        assert reader.file_path == Path("test.xlsx")
    
    @patch('pathlib.Path.exists')
    def test_init_with_path_object(self, mock_exists):
        """测试使用Path对象初始化"""
        mock_exists.return_value = True
        path = Path("test.xlsx")
        reader = ExcelReader(path)
        assert reader.file_path == path
    
    def test_init_file_not_found(self):
        """测试文件不存在时的异常"""
        with pytest.raises(FileNotFoundError):
            ExcelReader("nonexistent.xlsx")
    
    @patch('pathlib.Path.exists')
    def test_init_invalid_format(self, mock_exists):
        """测试无效文件格式"""
        mock_exists.return_value = True
        with pytest.raises(ValueError):
            ExcelReader("test.txt")
    
    @patch('pathlib.Path.exists')
    @patch('pandas.read_excel')
    def test_read_data_success(self, mock_read_excel, mock_exists):
        """测试成功读取数据"""
        mock_exists.return_value = True
        mock_df = pd.DataFrame({
            'name': ['张三', '李四'],
            'age': [25, 30],
            'city': ['北京', '上海']
        })
        mock_read_excel.return_value = mock_df
        
        reader = ExcelReader("test.xlsx")
        result = reader.read_data()
        
        assert result is not None
        assert isinstance(result, pd.DataFrame)
        assert len(result) == 2
        mock_read_excel.assert_called_once_with(reader.file_path, sheet_name=None)
    
    @patch('pathlib.Path.exists')
    @patch('pandas.read_excel')
    def test_read_data_with_sheet_name(self, mock_read_excel, mock_exists):
        """测试指定工作表读取数据"""
        mock_exists.return_value = True
        mock_df = pd.DataFrame({'name': ['张三']})
        mock_read_excel.return_value = mock_df
        
        reader = ExcelReader("test.xlsx")
        result = reader.read_data("Sheet1")
        
        mock_read_excel.assert_called_once_with(reader.file_path, sheet_name="Sheet1")
    
    @patch('pathlib.Path.exists')
    @patch('pandas.ExcelFile')
    def test_get_sheet_names(self, mock_excel_file, mock_exists):
        """测试获取工作表名称"""
        mock_exists.return_value = True
        mock_file = MagicMock()
        mock_file.sheet_names = ['Sheet1', 'Sheet2']
        mock_excel_file.return_value = mock_file
        
        reader = ExcelReader("test.xlsx")
        result = reader.get_sheet_names()
        
        assert result == ['Sheet1', 'Sheet2']
        mock_excel_file.assert_called_once_with(reader.file_path)