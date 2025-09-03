#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
基础步骤组件 - 所有步骤组件的基类
"""

import tkinter as tk
from tkinter import ttk
from abc import ABC, abstractmethod


class BaseStep(ABC):
    """步骤组件基类"""
    
    def __init__(self, parent, app_data, title="步骤"):
        """初始化步骤组件
        
        Args:
            parent: 父容器
            app_data: 应用数据对象
            title: 步骤标题
        """
        self.parent = parent
        self.app_data = app_data
        self.title = title
        
        # 创建主框架
        self.frame = ttk.Frame(parent)
        self.setup_ui()
    
    def setup_ui(self):
        """设置基础UI结构"""
        # 步骤标题
        title_label = tk.Label(
            self.frame, 
            text=self.title, 
            font=("Arial", 14, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # 内容区域
        self.content_frame = ttk.Frame(self.frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 让子类实现具体内容
        self.setup_content()
        
        # 错误信息标签
        self.error_label = tk.Label(
            self.frame, 
            text="", 
            foreground='red',
            font=("Arial", 10)
        )
        self.error_label.pack(pady=(10, 0))
    
    @abstractmethod
    def setup_content(self):
        """设置步骤具体内容 - 子类必须实现"""
        pass
    
    def on_show(self):
        """当步骤显示时调用 - 子类可以重写"""
        self.clear_error()
    
    @abstractmethod
    def validate(self) -> bool:
        """验证步骤数据 - 子类必须实现
        
        Returns:
            bool: 验证是否通过
        """
        pass
    
    @abstractmethod
    def save_data(self):
        """保存步骤数据到app_data - 子类必须实现"""
        pass
    
    def show_error(self, message):
        """显示错误信息
        
        Args:
            message: 错误信息
        """
        self.error_label.config(text=message)
    
    def clear_error(self):
        """清除错误信息"""
        self.error_label.config(text="")
    
    def create_label_entry_pair(self, parent, label_text, entry_var=None, **entry_kwargs):
        """创建标签-输入框对
        
        Args:
            parent: 父容器
            label_text: 标签文本
            entry_var: 输入框变量
            **entry_kwargs: 输入框额外参数
            
        Returns:
            tuple: (标签, 输入框)
        """
        container = ttk.Frame(parent)
        container.pack(fill=tk.X, pady=5)
        
        label = ttk.Label(container, text=label_text, width=20)
        label.pack(side=tk.LEFT)
        
        if entry_var is None:
            entry_var = tk.StringVar()
        
        entry = ttk.Entry(container, textvariable=entry_var, **entry_kwargs)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))
        
        return label, entry
    
    def create_label_combobox_pair(self, parent, label_text, values, combobox_var=None, **combobox_kwargs):
        """创建标签-下拉框对
        
        Args:
            parent: 父容器
            label_text: 标签文本
            values: 下拉框选项
            combobox_var: 下拉框变量
            **combobox_kwargs: 下拉框额外参数
            
        Returns:
            tuple: (标签, 下拉框)
        """
        container = ttk.Frame(parent)
        container.pack(fill=tk.X, pady=5)
        
        label = ttk.Label(container, text=label_text, width=20)
        label.pack(side=tk.LEFT)
        
        if combobox_var is None:
            combobox_var = tk.StringVar()
        
        combobox = ttk.Combobox(
            container, 
            textvariable=combobox_var, 
            values=values,
            state='readonly',
            **combobox_kwargs
        )
        combobox.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))
        
        return label, combobox
    
    def create_info_text(self, parent, text, **text_kwargs):
        """创建信息文本区域
        
        Args:
            parent: 父容器
            text: 文本内容
            **text_kwargs: 文本框额外参数
            
        Returns:
            tk.Text: 文本框组件
        """
        text_widget = tk.Text(
            parent, 
            height=8, 
            wrap=tk.WORD,
            state='disabled',
            **text_kwargs
        )
        text_widget.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 添加文本内容
        text_widget.config(state='normal')
        text_widget.insert(tk.END, text)
        text_widget.config(state='disabled')
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.config(yscrollcommand=scrollbar.set)
        
        return text_widget