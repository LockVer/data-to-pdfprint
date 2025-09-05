"""
通用Excel数据提取器
根据关键字动态查找并提取数据
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Any, Optional

class ExcelDataExtractor:
    """
    Excel数据提取器
    通过关键字查找对应的数据位置
    """
    
    def __init__(self, file_path: str):
        """
        初始化提取器
        
        Args:
            file_path: Excel文件路径
        """
        self.file_path = file_path
        self.df = None
        self.keyword_positions = {}
        self._load_excel()
    
    def _load_excel(self):
        """加载Excel文件"""
        try:
            self.df = pd.read_excel(self.file_path, header=None)
            print(f"✅ Excel文件已加载: {self.df.shape[0]}行 x {self.df.shape[1]}列")
        except Exception as e:
            raise Exception(f"无法加载Excel文件: {e}")
    
    def find_keyword(self, keyword: str) -> List[Dict[str, Any]]:
        """
        查找关键字在Excel中的位置
        
        Args:
            keyword: 要查找的关键字 
            
        Returns:
            包含位置信息的字典列表
        """
        if self.df is None:
            return []
        
        positions = []
        
        for row_idx in range(self.df.shape[0]):
            for col_idx in range(self.df.shape[1]):
                cell_value = self.df.iloc[row_idx, col_idx]
                
                if pd.notna(cell_value):
                    cell_str = str(cell_value).strip()
                    
                    if keyword in cell_str:
                        col_letter = self._col_index_to_letter(col_idx)
                        positions.append({
                            'row': row_idx,
                            'col': col_idx,
                            'excel_ref': f"{col_letter}{row_idx + 1}",
                            'value': cell_str,
                            'keyword': keyword
                        })
        
        return positions
    
    def _col_index_to_letter(self, col_idx: int) -> str:
        """
        将列索引转换为Excel列字母
        
        Args:
            col_idx: 列索引 (0-based)
            
        Returns:
            Excel列字母 (如 'A', 'B', 'AA')
        """
        result = ""
        while col_idx >= 0:
            result = chr(65 + col_idx % 26) + result
            col_idx = col_idx // 26 - 1
        return result
    
    def get_nearby_value(self, row: int, col: int, direction: str) -> Optional[Any]:
        """
        获取指定位置附近的值
        
        Args:
            row: 行索引
            col: 列索引  
            direction: 方向 ('right', 'down', 'left', 'up', 'right_down', etc.)
            
        Returns:
            附近单元格的值
        """
        if self.df is None:
            return None
        
        direction_map = {
            'right': (0, 1),
            'down': (1, 0),
            'left': (0, -1),
            'up': (-1, 0),
            'right_down': (1, 1),
            'left_down': (1, -1),
            'right_up': (-1, 1),
            'left_up': (-1, -1)
        }
        
        if direction not in direction_map:
            return None
        
        dr, dc = direction_map[direction]
        new_row, new_col = row + dr, col + dc
        
        # 检查边界
        if (0 <= new_row < self.df.shape[0] and 0 <= new_col < self.df.shape[1]):
            value = self.df.iloc[new_row, new_col]
            return value if pd.notna(value) else None
        
        return None
    
    def extract_data_by_keywords(self, keyword_config: Dict[str, Dict]) -> Dict[str, Any]:
        """
        根据关键字配置提取数据
        
        Args:
            keyword_config: 关键字配置字典
            格式: {
                'field_name': {
                    'keyword': '关键字',
                    'direction': 'right',  # 数据相对于关键字的位置
                    'offset': (0, 1)  # 可选，额外偏移
                }
            }
            
        Returns:
            提取的数据字典
        """
        extracted_data = {}
        
        for field_name, config in keyword_config.items():
            keyword = config['keyword']
            direction = config.get('direction', 'right')
            offset = config.get('offset', (0, 0))
            
            # 查找关键字
            positions = self.find_keyword(keyword)
            
            if positions:
                # 使用第一个匹配的位置
                pos = positions[0]
                row, col = pos['row'], pos['col']
                
                # 应用方向偏移
                if direction == 'right':
                    target_row, target_col = row, col + 1
                elif direction == 'down':
                    target_row, target_col = row + 1, col
                elif direction == 'left':
                    target_row, target_col = row, col - 1
                elif direction == 'up':
                    target_row, target_col = row - 1, col
                else:
                    # 直接使用关键字位置
                    target_row, target_col = row, col
                
                # 应用额外偏移
                target_row += offset[0]
                target_col += offset[1]
                
                # 获取目标位置的值
                if (0 <= target_row < self.df.shape[0] and 0 <= target_col < self.df.shape[1]):
                    value = self.df.iloc[target_row, target_col]
                    extracted_data[field_name] = value if pd.notna(value) else None
                    
                    col_letter = self._col_index_to_letter(target_col)
                    print(f"✅ {field_name}: 从 {col_letter}{target_row + 1} 提取 = {value}")
                else:
                    print(f"❌ {field_name}: 目标位置超出范围")
                    extracted_data[field_name] = None
            else:
                print(f"❌ {field_name}: 未找到关键字 '{keyword}'")
                extracted_data[field_name] = None
        
        return extracted_data
    
    def extract_common_data(self) -> Dict[str, Any]:
        """
        提取所有模板都需要的公共数据：客户编码、标签名称、开始号、总张数
        
        Returns:
            包含公共数据的字典
        """
        print("🔍 提取公共数据字段...")
        
        # 定义公共数据的关键字配置
        keyword_config = {
            '标签名称': {
                'keyword': '标签名称',
                'direction': 'right'
            },
            '开始号': {
                'keyword': '开始号', 
                'direction': 'down'
            },
            '客户编码': {
                'keyword': '客户名称编码',
                'direction': 'down'
            }
        }
        
        # 使用关键字提取数据
        extracted_data = self.extract_data_by_keywords(keyword_config)
        
        # 提取总张数（使用专门的逻辑）
        from src.utils.text_processor import text_processor
        total_count = text_processor.extract_total_count_by_keyword(self.df)
        extracted_data['总张数'] = total_count
        
        # 提取其他基础数据（兼容现有逻辑）
        try:
            if not extracted_data.get('客户编码'):
                # 备用：从固定位置提取客户编码
                extracted_data['客户编码'] = str(self.df.iloc[3, 0]) if pd.notna(self.df.iloc[3, 0]) else 'Unknown Client'
                
            if not extracted_data.get('标签名称'):
                # 备用：从固定位置提取主题作为标签名称
                extracted_data['标签名称'] = str(self.df.iloc[3, 1]) if pd.notna(self.df.iloc[3, 1]) else 'Unknown Title'
                
        except Exception as e:
            print(f"⚠️ 备用数据提取失败: {e}")
        
        print(f"✅ 公共数据提取完成:")
        for key, value in extracted_data.items():
            print(f"   {key}: {value}")
        
        return extracted_data

def test_extractor():
    """测试数据提取器"""
    file_path = '/Users/trq/Desktop/常规-LADIES NIGHT IN 女士夜._副本.xlsx'
    
    print("🚀 测试Excel数据提取器")
    print("=" * 50)
    
    try:
        extractor = ExcelDataExtractor(file_path)
        
        # 定义要提取的数据配置
        keyword_config = {
            '开始号': {
                'keyword': '开始号',
                'direction': 'down'  # 开始号在B10，数据在B11
            },
            '标签名称': {
                'keyword': '标签名称:',
                'direction': 'right'  # 标签名称在G11，数据在H11
            },
            '客户编码': {
                'keyword': '14KH0149',  # 直接查找客户编码
                'direction': 'none'
            },
            '主题': {
                'keyword': '主题',
                'direction': 'down'  # 主题标签在B3，数据在B4
            },
            '总张数': {
                'keyword': '总张数',
                'direction': 'down'  # 总张数标签，数据在下方
            }
        }
        
        # 提取数据
        print("\n📊 开始提取数据...")
        extracted_data = extractor.extract_data_by_keywords(keyword_config)
        
        print(f"\n📋 提取结果:")
        for field, value in extracted_data.items():
            print(f"  {field}: {value}")
        
        return extracted_data
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return None

if __name__ == "__main__":
    test_extractor()