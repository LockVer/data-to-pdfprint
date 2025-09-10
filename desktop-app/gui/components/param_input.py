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
        self.is_overweight_var = tk.BooleanVar(value=False)  # 超重模式选择
        
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
        
        # 每盒张数/每套张数（根据模式动态显示）
        self.sheets_container = self.create_dynamic_sheets_input(
            basic_frame,
            self.sheets_per_box_var
        )
        
        # 高级参数区域（根据模式显示）
        self.advanced_frame = ttk.LabelFrame(self.content_frame, text="箱标参数", padding=20)
        self.advanced_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 每小箱盒数/套数（根据模式动态显示）
        self.boxes_container = self.create_dynamic_param_input(
            self.advanced_frame,
            self.boxes_per_small_case_var
        )
        
        # 每大箱小箱数
        self.cases_container = self.create_param_input(
            self.advanced_frame,
            "每大箱小箱数:",
            self.small_cases_per_large_case_var,
            "每个大箱包含的小箱数量"
        )
        
        # 超重模式复选框（仅套盒模式显示）
        self.overweight_container = self.create_overweight_checkbox(self.advanced_frame)
        self.overweight_container.pack_forget()  # 默认隐藏，在套盒模式时显示
        
        # 计算预览区域
        preview_frame = ttk.LabelFrame(self.content_frame, text="参数预览", padding=15)
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        self.preview_text = tk.Text(
            preview_frame,
            height=8,
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
        self.is_overweight_var.trace_add('write', self.update_preview)
    
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
    
    def create_dynamic_sheets_input(self, parent, var):
        """创建动态张数标签的参数输入控件"""
        container = ttk.Frame(parent)
        container.pack(fill=tk.X, pady=10)
        
        # 标签（动态文本）
        self.sheets_label = tk.Label(container, text="每盒张数:", width=20, font=("Arial", 10, "bold"))
        self.sheets_label.pack(side=tk.LEFT)
        
        # 输入框
        entry = tk.Entry(container, textvariable=var, width=15, font=("Arial", 11))
        entry.pack(side=tk.LEFT, padx=(10, 0))
        
        # 帮助文本（动态文本）
        self.sheets_help_label = tk.Label(
            container, 
            text="每个盒子包含的卡片数量", 
            font=("Arial", 9),
            foreground='#666666'
        )
        self.sheets_help_label.pack(side=tk.LEFT, padx=(10, 0))
        
        return container
    
    def create_dynamic_param_input(self, parent, var):
        """创建动态标签的参数输入控件"""
        container = ttk.Frame(parent)
        container.pack(fill=tk.X, pady=10)
        
        # 标签（动态文本）
        self.boxes_label = tk.Label(container, text="每小箱盒数:", width=20, font=("Arial", 10, "bold"))
        self.boxes_label.pack(side=tk.LEFT)
        
        # 输入框
        self.boxes_entry = tk.Entry(container, textvariable=var, width=15, font=("Arial", 11))
        self.boxes_entry.pack(side=tk.LEFT, padx=(10, 0))
        
        # 帮助文本（动态文本）
        self.boxes_help_label = tk.Label(
            container, 
            text="每个小箱包含的盒子数量", 
            font=("Arial", 9),
            foreground='#666666'
        )
        self.boxes_help_label.pack(side=tk.LEFT, padx=(10, 0))
        
        return container
    
    def create_overweight_checkbox(self, parent):
        """创建超重模式复选框"""
        container = ttk.Frame(parent)
        
        # 复选框
        checkbox = tk.Checkbutton(
            container,
            text="超重模式",
            variable=self.is_overweight_var,
            font=("Arial", 10, "bold"),
            command=self.on_overweight_change
        )
        checkbox.pack(side=tk.LEFT)
        
        # 帮助文本
        help_label = tk.Label(
            container, 
            text="勾选表示一套需要分为多个箱子，将显示不同的参数设置", 
            font=("Arial", 9),
            foreground='#666666'
        )
        help_label.pack(side=tk.LEFT, padx=(10, 0))
        
        return container
    
    def on_overweight_change(self):
        """超重模式复选框变化时的回调"""
        if self.app_data.package_mode == 'set':
            self.update_cases_label_text()
            self.update_preview()
    
    def update_sheets_label_text(self):
        """根据包装模式更新基础参数标签文本"""
        if hasattr(self, 'sheets_label') and hasattr(self, 'sheets_help_label'):
            if self.app_data.package_mode == 'set':
                self.sheets_label.config(text="套中每盒张数:")
                self.sheets_help_label.config(text="套中每个盒子包含的卡片数量")
            else:
                self.sheets_label.config(text="每盒张数:")
                self.sheets_help_label.config(text="每个盒子包含的卡片数量")
    
    def update_boxes_label_text(self):
        """根据包装模式更新高级参数标签文本"""
        if hasattr(self, 'boxes_label') and hasattr(self, 'boxes_help_label'):
            if self.app_data.package_mode == 'set':
                self.boxes_label.config(text="每套盒数:")
                self.boxes_help_label.config(text="每个套包含的盒子数量")
            else:
                self.boxes_label.config(text="每小箱盒数:")
                self.boxes_help_label.config(text="每个小箱包含的盒子数量")
    
    def update_cases_label_text(self):
        """根据包装模式和超重状态更新箱数参数标签文本"""
        if hasattr(self, 'cases_container') and self.app_data.package_mode == 'set':
            # 获取cases_container中的标签组件
            for widget in self.cases_container.winfo_children():
                if isinstance(widget, tk.Label) and hasattr(widget, 'config'):
                    widget_text = widget.cget('text')
                    if '大箱' in widget_text or '套数' in widget_text or '分为' in widget_text:
                        if self.is_overweight_var.get():
                            widget.config(text="一套分为几箱:")
                        else:
                            widget.config(text="每大箱套数:")
                        break
            
            # 更新帮助文本
            for widget in self.cases_container.winfo_children():
                if isinstance(widget, tk.Label) and widget.cget('foreground') == '#666666':
                    if self.is_overweight_var.get():
                        widget.config(text="每个套分为几个箱子进行包装")
                    else:
                        widget.config(text="每个大箱包含的套数")
                    break
    
    def update_advanced_params_visibility(self):
        """根据包装模式更新高级参数可见性"""
        if hasattr(self, 'advanced_frame') and self.app_data.package_mode:
            if self.app_data.package_mode == 'regular':
                # 常规模式：显示"每小箱盒数"和"每大箱小箱数"，允许用户输入每小箱盒数
                self.boxes_container.pack(fill=tk.X, pady=10)
                self.cases_container.pack(fill=tk.X, pady=10)
                self.overweight_container.pack_forget()  # 隐藏超重模式复选框
                # 常规模式：恢复输入框为可编辑状态
                if hasattr(self, 'boxes_entry'):
                    self.boxes_entry.config(state='normal', bg='white')
                    self.boxes_help_label.config(text="每个小箱包含的盒子数量", foreground='#666666')
                # 设置默认值
                if self.boxes_per_small_case_var.get() <= 1:
                    self.boxes_per_small_case_var.set(1)  # 常规模式默认为1，但用户可以修改
                if self.small_cases_per_large_case_var.get() <= 1:
                    self.small_cases_per_large_case_var.set(2)  # 默认2个小箱一大箱
            elif self.app_data.package_mode == 'set':
                # 套盒模式：显示"每套盒数"、"每大箱套数/一套分为几箱"和超重模式复选框
                self.boxes_container.pack(fill=tk.X, pady=10)
                self.cases_container.pack(fill=tk.X, pady=10)
                self.overweight_container.pack(fill=tk.X, pady=10)  # 显示超重模式复选框
                # 套盒模式：恢复输入框为可编辑状态
                if hasattr(self, 'boxes_entry'):
                    self.boxes_entry.config(state='normal', bg='white')
                    self.boxes_help_label.config(text="每个套包含的盒子数量", foreground='#666666')
                # 设置默认值
                if self.boxes_per_small_case_var.get() <= 1:
                    self.boxes_per_small_case_var.set(6)  # 每套默认6盒
                if self.small_cases_per_large_case_var.get() <= 1:
                    self.small_cases_per_large_case_var.set(2)  # 每大箱默认2套（非超重）或一套分为2箱（超重）
                # 更新标签文本
                self.update_cases_label_text()
            else:
                # 分盒模式显示所有高级参数，允许用户输入每小箱盒数
                self.boxes_container.pack(fill=tk.X, pady=10)
                self.cases_container.pack(fill=tk.X, pady=10)
                self.overweight_container.pack_forget()  # 隐藏超重模式复选框
                # 分盒模式：允许用户输入每小箱盒数，不再写死为1
                if hasattr(self, 'boxes_entry'):
                    self.boxes_entry.config(state='normal', bg='white')
                    self.boxes_help_label.config(text="每个小箱包含的盒子数量", foreground='#666666')
                # 设置默认值
                if self.boxes_per_small_case_var.get() <= 0:
                    self.boxes_per_small_case_var.set(1)  # 默认为1，但用户可以修改
                if self.small_cases_per_large_case_var.get() <= 1:
                    self.small_cases_per_large_case_var.set(2)
        
        # 更新标签文本
        self.update_sheets_label_text()
        self.update_boxes_label_text()
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
                
                if self.app_data.package_mode == 'set':
                    # 套盒模式计算
                    cards_per_box_in_set = sheets_per_box  # 套中每盒张数（用户输入）
                    boxes_per_set = boxes_per_small_case  # 每套盒数（用户输入）
                    is_overweight = self.is_overweight_var.get()  # 是否超重模式
                    
                    cards_per_set = cards_per_box_in_set * boxes_per_set  # 每套张数（计算得出）
                    total_sets = math.ceil(total_sheets / cards_per_set) if cards_per_set > 0 else 0
                    total_boxes = total_sets * boxes_per_set
                    
                    if is_overweight:
                        # 超重模式：一套分为几箱
                        cases_per_set = small_cases_per_large_case  # 一套分为几箱（用户输入）
                        total_small_cases = total_sets * cases_per_set  # 总小箱数
                        total_large_cases = total_small_cases  # 超重模式下小箱就是最终包装
                    else:
                        # 非超重模式：每大箱套数
                        sets_per_large_case = small_cases_per_large_case  # 每大箱套数（用户输入）
                        total_small_cases = total_sets  # 一套=一小箱
                        total_large_cases = math.ceil(total_sets / sets_per_large_case) if sets_per_large_case > 0 else 0
                else:
                    # 其他模式原逻辑
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
                if self.app_data.package_mode == 'set':
                    preview_text += f"  套中每盒张数: {sheets_per_box:,} 张\n"
                    preview_text += f"  每套张数: {cards_per_set:,} 张（计算得出）\n"
                else:
                    preview_text += f"  每盒张数: {sheets_per_box:,} 张\n"
                
                if self.app_data.package_mode == 'regular':
                    # 常规模式：显示每小箱盒数和每大箱小箱数
                    preview_text += f"  每小箱盒数: {boxes_per_small_case} 盒\n"
                    preview_text += f"  每大箱小箱数: {small_cases_per_large_case} 小箱\n\n"
                elif self.app_data.package_mode == 'set':
                    # 套盒模式：显示用户输入的参数，根据超重模式显示不同内容
                    preview_text += f"  每套盒数: {boxes_per_small_case} 盒\n"
                    if is_overweight:
                        preview_text += f"  一套分为几箱: {small_cases_per_large_case} 箱（超重模式）\n\n"
                    else:
                        preview_text += f"  每大箱套数: {small_cases_per_large_case} 套（常规模式）\n\n"
                else:
                    # 分盒模式：显示所有参数（每小箱盒数可由用户输入）
                    preview_text += f"  每小箱盒数: {boxes_per_small_case} 盒\n"
                    preview_text += f"  每大箱小箱数: {small_cases_per_large_case} 小箱\n\n"
                
                preview_text += f"计算结果:\n"
                preview_text += f"  总盒数: {total_boxes} 盒\n"
                
                if self.app_data.package_mode == 'regular':
                    # 常规模式：显示小箱数(=总盒数)和大箱数
                    preview_text += f"  总小箱数: {total_small_cases} 小箱\n"
                    preview_text += f"  总大箱数: {total_large_cases} 大箱\n\n"
                elif self.app_data.package_mode == 'set':
                    # 套盒模式：根据超重模式显示不同结果
                    preview_text += f"  总套数: {total_sets} 套\n"
                    if is_overweight:
                        preview_text += f"  总小箱数: {total_small_cases} 小箱（超重模式，每套分{cases_per_set}箱）\n\n"
                    else:
                        preview_text += f"  总小箱数: {total_small_cases} 小箱（每套一箱）\n"
                        preview_text += f"  总大箱数: {total_large_cases} 大箱\n\n"
                else:
                    # 分盒模式：显示所有计算结果
                    preview_text += f"  总小箱数: {total_small_cases} 小箱\n"
                    preview_text += f"  总大箱数: {total_large_cases} 大箱\n\n"
                
                # 序列号示例
                preview_text += f"序列号示例:\n"
                
                if self.app_data.package_mode == 'regular':
                    preview_text += f"  单级格式: {start_number}, {self.get_next_serial(start_number)}, {self.get_next_serial(start_number, 2)}...\n"
                elif self.app_data.package_mode == 'set':
                    preview_text += f"  套盒格式: {start_number}-01到{start_number}-{boxes_per_small_case:02d}, {self.get_next_serial(start_number)}-01到{self.get_next_serial(start_number)}-{boxes_per_small_case:02d}...\n"
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
            
            # 常规/分盒/套盒模式需要额外验证"每小箱盒数"参数
            if self.app_data.package_mode == 'regular':
                # 常规模式：验证每小箱盒数参数
                if boxes_per_small_case <= 0:
                    self.show_error("每小箱盒数必须大于0")
                    return False
                
                if boxes_per_small_case > 100:
                    self.show_error("每小箱盒数不能超过100")
                    return False
            elif self.app_data.package_mode == 'separate':
                # 分盒模式：验证每小箱盒数参数（允许用户输入）
                if boxes_per_small_case <= 0:
                    self.show_error("每小箱盒数必须大于0")
                    return False
                
                if boxes_per_small_case > 100:
                    self.show_error("每小箱盒数不能超过100")
                    return False
            elif self.app_data.package_mode == 'set':
                # 套盒模式：验证每套盒数参数
                if boxes_per_small_case <= 0:
                    self.show_error("每套盒数必须大于0")
                    return False
                
                if boxes_per_small_case > 100:
                    self.show_error("每套盒数不能超过100")
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
            small_cases_per_large_case=self.small_cases_per_large_case_var.get(),
            is_overweight=self.is_overweight_var.get()
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
            self.is_overweight_var.set(params.is_overweight)
        
        self.update_preview()