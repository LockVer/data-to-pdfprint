#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
应用数据模型 - 管理整个应用的数据状态
"""

import os
from dataclasses import dataclass, field
from typing import Optional, Dict, Any


@dataclass
class PackageParams:
    """包装参数"""
    sheets_per_box: int = 2850
    boxes_per_small_case: int = 1
    small_cases_per_large_case: int = 2
    is_overweight: bool = False  # 是否超重模式 (仅套盒模式使用)


@dataclass
class ExcelData:
    """从Excel提取的数据"""
    customer_code: str = ""
    theme: str = ""
    start_number: str = ""
    total_sheets: int = 0
    file_path: str = ""


class AppData:
    """应用数据管理类"""
    
    def __init__(self):
        """初始化应用数据"""
        # 文件相关
        self.excel_file_path: Optional[str] = None
        self.excel_data: Optional[ExcelData] = None
        
        # 包装模式: 'regular', 'separate', 'set'
        self.package_mode: Optional[str] = None
        
        # 包装参数
        self.package_params: PackageParams = PackageParams()
        
        # 标签模板: 'regular', 'game'
        self.label_template: Optional[str] = None
        
        # 输出相关
        self.output_directory: Optional[str] = None
        self.generated_files: list = field(default_factory=list)
        
        # 状态标志
        self.is_generating: bool = False
        self.generation_error: Optional[str] = None
    
    def reset(self):
        """重置所有数据"""
        self.__init__()
    
    def is_excel_file_valid(self) -> bool:
        """检查Excel文件是否有效"""
        return (self.excel_file_path is not None and 
                os.path.exists(self.excel_file_path) and
                self.excel_file_path.endswith(('.xlsx', '.xls')))
    
    def is_package_mode_valid(self) -> bool:
        """检查包装模式是否有效"""
        return self.package_mode in ['regular', 'separate', 'set']
    
    def is_package_params_valid(self) -> bool:
        """检查包装参数是否有效"""
        if not self.package_params:
            return False
        
        # 基本参数验证
        if self.package_params.sheets_per_box <= 0:
            return False
        
        # 根据模式验证参数
        if self.package_mode in ['separate', 'set']:
            if (self.package_params.boxes_per_small_case <= 0 or
                self.package_params.small_cases_per_large_case <= 0):
                return False
        
        return True
    
    def is_label_template_valid(self) -> bool:
        """检查标签模板是否有效"""
        return self.label_template in ['regular', 'game']
    
    def get_validation_errors(self) -> list:
        """获取所有验证错误"""
        errors = []
        
        if not self.is_excel_file_valid():
            errors.append("请选择有效的Excel文件")
        
        if not self.is_package_mode_valid():
            errors.append("请选择包装模式")
        
        if not self.is_package_params_valid():
            errors.append("请输入有效的包装参数")
        
        if not self.is_label_template_valid():
            errors.append("请选择标签模板")
        
        return errors
    
    def is_ready_for_generation(self) -> bool:
        """检查是否准备好生成PDF"""
        return len(self.get_validation_errors()) == 0
    
    def get_package_mode_display_name(self) -> str:
        """获取包装模式的显示名称"""
        mode_names = {
            'regular': '常规模式（单级）',
            'separate': '分盒模式（多级）',
            'set': '套盒模式（多级）'
        }
        return mode_names.get(self.package_mode, '未知模式')
    
    def get_label_template_display_name(self) -> str:
        """获取标签模板的显示名称"""
        template_names = {
            'regular': 'Regular模板（客户编码+主题+序列号）',
            'game': 'Game模板（Game title + Ticket count + Serial）'
        }
        return template_names.get(self.label_template, '未知模板')
    
    def get_output_file_names(self) -> tuple:
        """生成输出文件名"""
        if not self.excel_data:
            return "output_盒标.pdf", "output_箱标.pdf"
        
        # 基于Excel数据生成文件名
        customer = self.excel_data.customer_code or "客户"
        theme = self.excel_data.theme or "订单"
        
        folder_name = f"{customer}+{theme}+标签"
        box_label_name = f"{customer}+{theme}+盒标.pdf"
        case_label_name = f"{customer}+{theme}+箱标.pdf"
        
        return folder_name, box_label_name, case_label_name
    
    def get_summary(self) -> Dict[str, Any]:
        """获取配置摘要"""
        folder_name, box_name, case_name = self.get_output_file_names()
        
        return {
            'excel_file': os.path.basename(self.excel_file_path) if self.excel_file_path else "未选择",
            'package_mode': self.get_package_mode_display_name(),
            'label_template': self.get_label_template_display_name(),
            'sheets_per_box': self.package_params.sheets_per_box,
            'boxes_per_small_case': self.package_params.boxes_per_small_case if self.package_mode in ['separate', 'set'] else None,
            'small_cases_per_large_case': self.package_params.small_cases_per_large_case if self.package_mode in ['separate', 'set'] else None,
            'output_folder': folder_name,
            'box_label_file': box_name,
            'case_label_file': case_name
        }