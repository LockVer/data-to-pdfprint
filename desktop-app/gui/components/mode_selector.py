#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
包装模式选择步骤组件
"""

import tkinter as tk
from tkinter import ttk

from gui.components.base_step import BaseStep


class ModeSelectorStep(BaseStep):
    """包装模式选择步骤"""
    
    def __init__(self, parent, app_data, title="选择包装模式"):
        self.mode_var = tk.StringVar()
        self.mode_frames = {}
        super().__init__(parent, app_data, title)
    
    def setup_content(self):
        """设置步骤具体内容"""
        # 说明文本
        info_label = tk.Label(
            self.content_frame, 
            text="请选择适合的包装模式，这将影响序列号格式和箱标样式",
            font=("Arial", 11)
        )
        info_label.pack(pady=(0, 20))
        
        # 模式选择区域
        modes_frame = ttk.LabelFrame(self.content_frame, text="包装模式", padding=20)
        modes_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 模式选项
        modes = [
            {
                'value': 'regular',
                'title': '常规模式（单级）',
                'description': '适用于简单包装，使用单级序列号\n序列号格式：JAW01001, JAW01002...\n箱标类型：常规单级箱标',
                'example': '示例：每盒2850张，序列号 LAN01001, LAN01002, LAN01003...'
            },
            {
                'value': 'separate',
                'title': '分盒模式（多级）', 
                'description': '适用于分盒包装，使用多级序列号\n序列号格式：JAW01001-01, JAW01001-02...\n箱标类型：多级分盒箱标',
                'example': '示例：每盒2850张，每小箱4盒\n序列号 LAN01001-01, LAN01001-02, LAN01002-01...'
            },
            {
                'value': 'set', 
                'title': '套盒模式（多级）',
                'description': '适用于套装包装，使用多级序列号\n序列号格式：JAW01001-01, JAW01001-02...\n箱标类型：多级分套箱标',
                'example': '示例：每套6盒，每小箱1套\n序列号 JAW01001-01, JAW01001-02, JAW01001-03...'
            }
        ]
        
        for i, mode in enumerate(modes):
            # 单选按钮和标题
            mode_frame = ttk.Frame(modes_frame)
            mode_frame.pack(fill=tk.X, pady=(0, 15))
            
            # 单选按钮
            radio_btn = ttk.Radiobutton(
                mode_frame,
                text=mode['title'],
                variable=self.mode_var,
                value=mode['value'],
                command=self.on_mode_change
            )
            radio_btn.pack(anchor=tk.W)
            
            # 描述信息
            desc_frame = ttk.Frame(mode_frame)
            desc_frame.pack(fill=tk.X, padx=(20, 0), pady=(5, 0))
            
            desc_label = ttk.Label(
                desc_frame,
                text=mode['description']
            )
            desc_label.pack(anchor=tk.W)
            
            # 示例信息
            example_label = ttk.Label(
                desc_frame,
                text=mode['example']
            )
            example_label.pack(anchor=tk.W, pady=(3, 0))
            
            self.mode_frames[mode['value']] = mode_frame
        
        # 详细说明区域
        detail_frame = ttk.LabelFrame(self.content_frame, text="模式详细说明", padding=15)
        detail_frame.pack(fill=tk.BOTH, expand=True)
        
        self.detail_text = tk.Text(
            detail_frame,
            height=8,
            wrap=tk.WORD,
            state='disabled',
            font=("Arial", 10),
            bg='#f8f8f8'
        )
        self.detail_text.pack(fill=tk.BOTH, expand=True)
        
        # 绑定变量变化事件
        self.mode_var.trace_add('write', self.update_detail)
        
        # 默认选择第一个模式
        self.mode_var.set('regular')
        self.update_detail()
    
    def on_mode_change(self):
        """模式变化处理"""
        self.clear_error()
        self.update_detail()
    
    def update_detail(self, *args):
        """更新详细说明"""
        mode = self.mode_var.get()
        
        details = {
            'regular': """常规模式（单级编号）

特点：
• 使用单级序列号，格式简单直观
• 每个盒子有独立的序列号（如 LAN01001, LAN01002）
• 适合传统包装方式
• 箱标使用常规单级样式

适用场景：
• 简单包装需求
• 不需要复杂分级管理
• 传统游戏卡片包装

参数要求：
• 只需要设置每盒张数
• 每小箱盒数自动设为1
• 生成的序列号连续递增""",

            'separate': """分盒模式（多级编号）

特点：
• 使用多级序列号，支持分盒管理
• 序列号格式：主号-子号（如 LAN01001-01, LAN01001-02）
• 每个小箱包含多个盒子
• 箱标使用多级分盒样式

适用场景：
• 需要分盒包装管理
• 每个主序列号下有多个子包装
• 便于库存分类管理

参数要求：
• 需要设置每盒张数
• 需要设置每小箱盒数（大于1）
• 需要设置每大箱小箱数
• 序列号按分盒规则生成""",

            'set': """套盒模式（多级编号）

特点：
• 使用多级序列号，支持套装管理
• 序列号格式：主号-套号（如 JAW01001-01, JAW01001-02）
• 每套包含固定数量的盒子
• 箱标使用多级分套样式

适用场景：
• 套装产品包装
• 每个主题下有多个套装
• 需要按套装分类管理

参数要求：
• 需要设置每盒张数
• 需要设置每小箱盒数（套装内盒数）
• 需要设置每大箱小箱数
• 序列号按套装规则生成"""
        }
        
        detail_text = details.get(mode, "")
        
        self.detail_text.config(state='normal')
        self.detail_text.delete(1.0, tk.END)
        self.detail_text.insert(tk.END, detail_text)
        self.detail_text.config(state='disabled')
    
    def validate(self) -> bool:
        """验证模式选择"""
        mode = self.mode_var.get()
        
        if not mode:
            self.show_error("请选择包装模式")
            return False
        
        if mode not in ['regular', 'separate', 'set']:
            self.show_error("选择的包装模式无效")
            return False
        
        return True
    
    def save_data(self):
        """保存模式选择到app_data"""
        self.app_data.package_mode = self.mode_var.get()
    
    def on_show(self):
        """当步骤显示时调用"""
        super().on_show()
        
        # 如果已经有数据，恢复选择
        if self.app_data.package_mode:
            self.mode_var.set(self.app_data.package_mode)
            self.update_detail()