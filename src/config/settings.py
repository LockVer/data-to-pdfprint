"""
项目配置管理

管理项目的各种配置选项
"""

import os
from pathlib import Path

class Settings:
    """
    设置管理类
    """
    
    def __init__(self):
        """
        初始化设置
        """
        # 项目根目录
        self.PROJECT_ROOT = Path(__file__).parent.parent.parent
        
        # 模板目录
        self.TEMPLATE_DIR = self.PROJECT_ROOT / "templates"
        
        # 默认输出目录
        self.OUTPUT_DIR = self.PROJECT_ROOT / "output"
    
    def load_config(self, config_file: str = None):
        """
        加载配置文件
        
        Args:
            config_file: 配置文件路径
        """
        pass
    
    def save_config(self, config_file: str = None):
        """
        保存配置文件
        
        Args:
            config_file: 配置文件路径
        """
        pass

# 全局设置实例
settings = Settings()