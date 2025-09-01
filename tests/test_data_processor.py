"""
数据处理器测试
"""

import pytest
import pandas as pd
from unittest.mock import Mock
from src.data.data_processor import DataProcessor


class TestDataProcessor:
    """DataProcessor类的测试用例"""
    
    def test_init(self):
        """测试初始化"""
        processor = DataProcessor()
        assert processor is not None
    
    def test_extract_fields_with_valid_data(self):
        """测试从有效数据中提取字段"""
        processor = DataProcessor()
        
        # 创建测试数据
        raw_data = pd.DataFrame({
            'name': ['张三', '李四', '王五'],
            'age': [25, 30, 28],
            'city': ['北京', '上海', '深圳'],
            'extra': ['extra1', 'extra2', 'extra3']
        })
        
        # 测试提取指定字段
        fields = ['name', 'age', 'city']
        result = processor.extract_fields(raw_data, fields)
        
        assert result is not None
        assert isinstance(result, pd.DataFrame)
        assert list(result.columns) == fields
        assert len(result) == 3
    
    def test_extract_fields_missing_columns(self):
        """测试提取不存在的字段"""
        processor = DataProcessor()
        
        raw_data = pd.DataFrame({
            'name': ['张三', '李四'],
            'age': [25, 30]
        })
        
        fields = ['name', 'nonexistent']
        with pytest.raises(KeyError):
            processor.extract_fields(raw_data, fields)
    
    def test_validate_data_valid(self):
        """测试验证有效数据"""
        processor = DataProcessor()
        
        valid_data = pd.DataFrame({
            'name': ['张三', '李四'],
            'age': [25, 30],
            'city': ['北京', '上海']
        })
        
        result = processor.validate_data(valid_data)
        assert result is True
    
    def test_validate_data_empty(self):
        """测试验证空数据"""
        processor = DataProcessor()
        
        empty_data = pd.DataFrame()
        result = processor.validate_data(empty_data)
        assert result is False
    
    def test_validate_data_none(self):
        """测试验证None数据"""
        processor = DataProcessor()
        
        result = processor.validate_data(None)
        assert result is False
    
    def test_validate_data_with_null_values(self):
        """测试包含空值的数据验证"""
        processor = DataProcessor()
        
        data_with_nulls = pd.DataFrame({
            'name': ['张三', None, '王五'],
            'age': [25, 30, None],
            'city': ['北京', '上海', '深圳']
        })
        
        result = processor.validate_data(data_with_nulls)
        assert result is False
    
    def test_clean_data(self):
        """测试数据清理功能"""
        processor = DataProcessor()
        
        dirty_data = pd.DataFrame({
            'name': ['  张三  ', 'Li四', '王五'],
            'age': ['25', '30.0', '28'],
            'city': ['北京', '上海', '深圳']
        })
        
        result = processor.clean_data(dirty_data)
        
        assert result is not None
        assert isinstance(result, pd.DataFrame)
        # 测试数据清理后的效果
        assert result.iloc[0]['name'] == '张三'  # 去除空格
        # 修复类型断言 - pandas Int64 不等同于 Python int
        assert pd.api.types.is_integer_dtype(result['age'])  # 类型转换