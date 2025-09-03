#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
包装参数输入步骤组件
"""

import tkinter as tk
from tkinter import ttk

from gui.components.base_step import BaseStep
from core.app_data import PackageParams


class ParamInputStep(BaseStep):
    """包装参数输入步骤"""
    
    def __init__(self, parent, app_data, title="输入包装参数"):
        self.sheets_per_box_var = tk.IntVar(value=2850)
        self.boxes_per_small_case_var = tk.IntVar(value=4)
        self.small_cases_per_large_case_var = tk.IntVar(value=2)
        
        self.advanced_frame = None
        super().__init__(parent, app_data, title)
    
    def setup_content(self):
        """设置步骤具体内容"""
        # 说明文本
        info_label = tk.Label(
            self.content_frame, 
            text="请根据选择的包装模式输入相应的包装参数",
            font=("Arial", 11)
        )
        info_label.pack(pady=(0, 20))
        
        # 基础参数区域
        basic_frame = ttk.LabelFrame(self.content_frame, text="基础参数", padding=20)
        basic_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 每盒张数
        self.create_param_input(
            basic_frame,
            "每盒张数:",
            self.sheets_per_box_var,
            "每个盒子包含的卡片数量"
        )
        
        # 高级参数区域（根据模式显示）
        self.advanced_frame = ttk.LabelFrame(self.content_frame, text="箱标参数", padding=20)
        self.advanced_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 每小箱盒数
        self.boxes_container = self.create_param_input(
            self.advanced_frame,
            "每小箱盒数:",
            self.boxes_per_small_case_var,
            "每个小箱包含的盒子数量"
        )
        
        # 每大箱小箱数
        self.cases_container = self.create_param_input(
            self.advanced_frame,
            "每大箱小箱数:",
            self.small_cases_per_large_case_var,
            "每个大箱包含的小箱数量"
        )
        
        # 计算预览区域
        preview_frame = ttk.LabelFrame(self.content_frame, text="参数预览", padding=15)
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        self.preview_text = tk.Text(
            preview_frame,
            height=10,
            wrap=tk.WORD,
            state='disabled',
            font=("Arial", 10),
            bg='#f8f8f8'
        )
        self.preview_text.pack(fill=tk.BOTH, expand=True)
        
        # 绑定变量变化事件
        self.sheets_per_box_var.trace_add('write', self.update_preview)
        self.boxes_per_small_case_var.trace_add('write', self.update_preview)
        self.small_cases_per_large_case_var.trace_add('write', self.update_preview)
    
    def create_param_input(self, parent, label_text, var, help_text):
        """创建参数输入控件"""
        container = ttk.Frame(parent)
        container.pack(fill=tk.X, pady=10)
        
        # 标签
        label = tk.Label(container, text=label_text, width=20, font=("Arial", 10, "bold"))
        label.pack(side=tk.LEFT)
        
        # 输入框
        entry = tk.Entry(container, textvariable=var, width=15, font=("Arial", 11))
        entry.pack(side=tk.LEFT, padx=(10, 0))
        
        # 帮助文本
        help_label = tk.Label(
            container, 
            text=help_text, 
            font=("Arial", 9),
            foreground='#666666'
        )
        help_label.pack(side=tk.LEFT, padx=(10, 0))
        
        return container
    
    def update_advanced_params_visibility(self):
        """根据包装模式更新高级参数可见性"""
        if hasattr(self, 'advanced_frame') and self.app_data.package_mode:
            if self.app_data.package_mode == 'regular':
                # 常规模式：隐藏"每小箱盒数"，显示"每大箱小箱数"
                self.boxes_container.pack_forget()
                self.cases_container.pack(fill=tk.X, pady=10)
                # 自动设置默认值
                self.boxes_per_small_case_var.set(1)  # 常规模式固定为1
                if self.small_cases_per_large_case_var.get() <= 1:
                    self.small_cases_per_large_case_var.set(2)  # 默认2个小箱一大箱
            else:
                # 分盒/套盒模式显示所有高级参数
                self.boxes_container.pack(fill=tk.X, pady=10)
                self.cases_container.pack(fill=tk.X, pady=10)
                # 设置默认值
                if self.boxes_per_small_case_var.get() <= 1:
                    self.boxes_per_small_case_var.set(4)
                if self.small_cases_per_large_case_var.get() <= 1:
                    self.small_cases_per_large_case_var.set(2)
        
        self.update_preview()
    
    def update_preview(self, *args):
        """更新参数预览"""
        try:
            sheets_per_box = self.sheets_per_box_var.get()
            boxes_per_small_case = self.boxes_per_small_case_var.get()
            small_cases_per_large_case = self.small_cases_per_large_case_var.get()
            
            if not self.app_data.excel_data:
                preview_text = "请先上传Excel文件以查看详细预览"
            else:
                total_sheets = self.app_data.excel_data.total_sheets
                theme = self.app_data.excel_data.theme
                start_number = self.app_data.excel_data.start_number
                
                # 计算相关数量
                import math
                total_boxes = math.ceil(total_sheets / sheets_per_box) if sheets_per_box > 0 else 0
                total_small_cases = math.ceil(total_boxes / boxes_per_small_case) if boxes_per_small_case > 0 else 0
                total_large_cases = math.ceil(total_small_cases / small_cases_per_large_case) if small_cases_per_large_case > 0 else 0
                
                preview_text = f"参数配置预览\n"
                preview_text += f"{'='*50}\n\n"
                
                preview_text += f"Excel数据:\n"
                preview_text += f"  主题: {theme}\n"
                preview_text += f"  开始号: {start_number}\n"
                preview_text += f"  总卡片数: {total_sheets:,} 张\n\n"
                
                preview_text += f"包装参数:\n"
                preview_text += f"  每盒张数: {sheets_per_box:,} 张\n"
                
                if self.app_data.package_mode == 'regular':
                    # 常规模式：只显示每大箱小箱数
                    preview_text += f"  每大箱小箱数: {small_cases_per_large_case} 小箱\n\n"
                else:
                    # 分盒/套盒模式：显示所有参数
                    preview_text += f"  每小箱盒数: {boxes_per_small_case} 盒\n"
                    preview_text += f"  每大箱小箱数: {small_cases_per_large_case} 小箱\n\n"
                
                preview_text += f"计算结果:\n"
                preview_text += f"  总盒数: {total_boxes} 盒\n"
                
                if self.app_data.package_mode == 'regular':
                    # 常规模式：显示小箱数(=总盒数)和大箱数
                    preview_text += f"  总小箱数: {total_small_cases} 小箱\n"
                    preview_text += f"  总大箱数: {total_large_cases} 大箱\n\n"
                else:
                    # 分盒/套盒模式：显示所有计算结果
                    preview_text += f"  总小箱数: {total_small_cases} 小箱\n"
                    preview_text += f"  总大箱数: {total_large_cases} 大箱\n\n"
                
                # 序列号示例
                preview_text += f"序列号示例:\n"
                
                if self.app_data.package_mode == 'regular':
                    preview_text += f"  单级格式: {start_number}, {self.get_next_serial(start_number)}, {self.get_next_serial(start_number, 2)}...\n"
                else:
                    preview_text += f"  多级格式: {start_number}-01, {start_number}-02, {self.get_next_serial(start_number)}-01...\n"
                
                preview_text += f"\n生成的PDF将包含:\n"
                preview_text += f"  • 盒标: {total_boxes} 页\n"
                
                if self.app_data.package_mode == 'regular':
                    preview_text += f"  • 箱标: {total_small_cases + total_large_cases} 页（常规单级：{total_small_cases}小箱+{total_large_cases}大箱）\n"
                elif self.app_data.package_mode == 'separate':
                    preview_text += f"  • 箱标: {total_small_cases + total_large_cases} 页（分盒多级）\n"
                else:  # set
                    preview_text += f"  • 箱标: {total_small_cases + total_large_cases} 页（分套多级）\n"
                
        except Exception as e:
            preview_text = f"参数预览计算出错: {str(e)}\n\n请检查输入的参数是否为有效数字"
        
        self.preview_text.config(state='normal')
        self.preview_text.delete(1.0, tk.END)
        self.preview_text.insert(tk.END, preview_text)
        self.preview_text.config(state='disabled')
    
    def get_next_serial(self, start_number, offset=1):
        """生成下一个序列号示例"""
        try:
            import re
            match = re.match(r'([A-Za-z]+)(\d+)', start_number)
            if match:
                prefix = match.group(1)
                number = int(match.group(2)) + offset
                return f"{prefix}{number:05d}"
            else:
                return f"{start_number}+{offset}"
        except:
            return f"{start_number}+{offset}"
    
    def validate(self) -> bool:
        """验证参数输入"""
        try:
            sheets_per_box = self.sheets_per_box_var.get()
            boxes_per_small_case = self.boxes_per_small_case_var.get()
            small_cases_per_large_case = self.small_cases_per_large_case_var.get()
            
            # 基础参数验证
            if sheets_per_box <= 0:
                self.show_error("每盒张数必须大于0")
                return False
            
            if sheets_per_box > 100000:
                self.show_error("每盒张数不能超过100,000")
                return False
            
            # 所有模式都需要验证"每大箱小箱数"参数
            if small_cases_per_large_case <= 0:
                self.show_error("每大箱小箱数必须大于0")
                return False
            
            if small_cases_per_large_case > 50:
                self.show_error("每大箱小箱数不能超过50")
                return False
            
            # 分盒/套盒模式需要额外验证"每小箱盒数"参数
            if self.app_data.package_mode in ['separate', 'set']:
                if boxes_per_small_case <= 0:
                    self.show_error("每小箱盒数必须大于0")
                    return False
                
                if boxes_per_small_case > 100:
                    self.show_error("每小箱盒数不能超过100")
                    return False
            
            return True
            
        except tk.TclError:
            self.show_error("请输入有效的数字")
            return False
    
    def save_data(self):
        """保存参数数据到app_data"""
        self.app_data.package_params = PackageParams(
            sheets_per_box=self.sheets_per_box_var.get(),
            boxes_per_small_case=self.boxes_per_small_case_var.get(),
            small_cases_per_large_case=self.small_cases_per_large_case_var.get()
        )
    
    def on_show(self):
        """当步骤显示时调用"""
        super().on_show()
        
        # 根据包装模式更新界面
        self.update_advanced_params_visibility()
        
        # 如果已经有数据，恢复设置
        if self.app_data.package_params:
            params = self.app_data.package_params
            self.sheets_per_box_var.set(params.sheets_per_box)
            self.boxes_per_small_case_var.set(params.boxes_per_small_case)
            self.small_cases_per_large_case_var.set(params.small_cases_per_large_case)
        
        self.update_preview()