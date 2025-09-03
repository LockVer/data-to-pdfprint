#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标签模板选择步骤组件
"""

import tkinter as tk
from tkinter import ttk

from gui.components.base_step import BaseStep


class TemplateSelectorStep(BaseStep):
    """标签模板选择步骤"""
    
    def __init__(self, parent, app_data, title="选择标签模板"):
        self.template_var = tk.StringVar()
        super().__init__(parent, app_data, title)
    
    def setup_content(self):
        """设置步骤具体内容"""
        # 说明文本
        info_label = tk.Label(
            self.content_frame, 
            text="请选择盒标的呈现效果，箱标将根据包装模式自动配套生成",
            font=("Arial", 11)
        )
        info_label.pack(pady=(0, 20))
        
        # 模板选择区域
        templates_frame = ttk.LabelFrame(self.content_frame, text="盒标模板", padding=20)
        templates_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 模板选项
        templates = [
            {
                'value': 'regular',
                'title': 'Regular模板',
                'description': '传统格式，适合常规业务场景',
                'content': '首页：客户编码 + 主题\n其他页：主题 + 序列号',
                'example': '示例页面：\n  14KH0095\n  JAWS\n\n  JAWS\n  JAW01001'
            },
            {
                'value': 'game',
                'title': 'Game模板',
                'description': '游戏信息格式，突出游戏元素',
                'content': '每页：Game title + Ticket count + Serial number',
                'example': '示例页面：\n  Game title: JAWS\n  Ticket count: 2850\n  Serial: JAW01001'
            }
        ]
        
        for template in templates:
            # 模板选择容器
            template_frame = ttk.Frame(templates_frame)
            template_frame.pack(fill=tk.X, pady=(0, 20))
            
            # 单选按钮
            radio_btn = ttk.Radiobutton(
                template_frame,
                text=template['title'],
                variable=self.template_var,
                value=template['value'],
                command=self.on_template_change
            )
            radio_btn.pack(anchor=tk.W)
            
            # 描述和内容
            desc_frame = ttk.Frame(template_frame)
            desc_frame.pack(fill=tk.X, padx=(20, 0), pady=(8, 0))
            
            # 左侧：描述和内容格式
            left_frame = ttk.Frame(desc_frame)
            left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            desc_label = tk.Label(
                left_frame,
                text=template['description'],
                font=("Arial", 10),
                foreground='#666666'
            )
            desc_label.pack(anchor=tk.W)
            
            content_label = tk.Label(
                left_frame,
                text=template['content'],
                font=("Arial", 10),
                foreground='#444444'
            )
            content_label.pack(anchor=tk.W, pady=(5, 0))
            
            # 右侧：示例预览
            right_frame = ttk.Frame(desc_frame)
            right_frame.pack(side=tk.RIGHT, padx=(20, 0))
            
            example_frame = ttk.LabelFrame(right_frame, text="预览", padding=10)
            example_frame.pack()
            
            example_label = tk.Label(
                example_frame,
                text=template['example'],
                font=("Courier", 9),
                foreground='#333333',
                background='#f8f8f8',
                relief='solid',
                borderwidth=1
            )
            example_label.pack()
        
        # 配置摘要区域
        summary_frame = ttk.LabelFrame(self.content_frame, text="配置摘要", padding=15)
        summary_frame.pack(fill=tk.BOTH, expand=True)
        
        self.summary_text = tk.Text(
            summary_frame,
            height=12,
            wrap=tk.WORD,
            state='disabled',
            font=("Arial", 10),
            bg='#f8f8f8'
        )
        self.summary_text.pack(fill=tk.BOTH, expand=True)
        
        # 绑定变量变化事件
        self.template_var.trace_add('write', self.update_summary)
        
        # 默认选择第一个模板
        self.template_var.set('regular')
        self.update_summary()
    
    def on_template_change(self):
        """模板变化处理"""
        self.clear_error()
        self.update_summary()
    
    def update_summary(self, *args):
        """更新配置摘要"""
        try:
            summary = self.app_data.get_summary()
            template = self.template_var.get()
            
            summary_text = f"最终配置摘要\n"
            summary_text += f"{'='*50}\n\n"
            
            # Excel文件信息
            summary_text += f"数据源:\n"
            summary_text += f"  文件: {summary['excel_file']}\n\n"
            
            # 包装配置
            summary_text += f"包装配置:\n"
            summary_text += f"  模式: {summary['package_mode']}\n"
            summary_text += f"  每盒张数: {summary['sheets_per_box']:,} 张\n"
            
            if summary.get('boxes_per_small_case'):
                summary_text += f"  每小箱盒数: {summary['boxes_per_small_case']} 盒\n"
            
            if summary.get('small_cases_per_large_case'):
                summary_text += f"  每大箱小箱数: {summary['small_cases_per_large_case']} 小箱\n"
            
            summary_text += f"\n"
            
            # 模板配置
            template_name = summary['label_template'] if template else "Regular模板（客户编码+主题+序列号）"
            summary_text += f"标签配置:\n"
            summary_text += f"  盒标模板: {template_name}\n"
            
            # 自动确定箱标模板
            mode = self.app_data.package_mode
            if mode == 'regular':
                case_template = "常规单级箱标"
            elif mode == 'separate':
                case_template = "多级分盒箱标"
            elif mode == 'set':
                case_template = "多级分套箱标"
            else:
                case_template = "自动选择"
            
            summary_text += f"  箱标模板: {case_template} (自动匹配)\n\n"
            
            # 输出文件
            summary_text += f"输出文件:\n"
            summary_text += f"  文件夹: {summary['output_folder']}/\n"
            summary_text += f"    ├── {summary['box_label_file']}\n"
            summary_text += f"    └── {summary['case_label_file']}\n\n"
            
            summary_text += f"输出格式: PDF/X (CMYK打印格式)\n"
            summary_text += f"页面尺寸: 90mm × 50mm\n\n"
            
            # 生成提示
            if self.app_data.is_ready_for_generation():
                summary_text += f"✅ 配置完成，可以生成PDF文件"
            else:
                errors = self.app_data.get_validation_errors()
                summary_text += f"❌ 配置不完整:\n"
                for error in errors:
                    summary_text += f"   • {error}\n"
            
        except Exception as e:
            summary_text = f"配置摘要生成出错: {str(e)}\n\n请检查之前的配置步骤"
        
        self.summary_text.config(state='normal')
        self.summary_text.delete(1.0, tk.END)
        self.summary_text.insert(tk.END, summary_text)
        self.summary_text.config(state='disabled')
    
    def validate(self) -> bool:
        """验证模板选择"""
        template = self.template_var.get()
        
        if not template:
            self.show_error("请选择标签模板")
            return False
        
        if template not in ['regular', 'game']:
            self.show_error("选择的标签模板无效")
            return False
        
        return True
    
    def save_data(self):
        """保存模板选择到app_data"""
        self.app_data.label_template = self.template_var.get()
    
    def on_show(self):
        """当步骤显示时调用"""
        super().on_show()
        
        # 更新摘要显示
        self.update_summary()
        
        # 如果已经有数据，恢复选择
        if self.app_data.label_template:
            self.template_var.set(self.app_data.label_template)
            self.update_summary()