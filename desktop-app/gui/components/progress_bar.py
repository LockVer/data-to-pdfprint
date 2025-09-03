#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
进度条组件 - 显示当前步骤进度
"""

import tkinter as tk
from tkinter import ttk


class ProgressBar:
    """步骤进度条组件"""
    
    def __init__(self, parent, total_steps=5):
        """初始化进度条
        
        Args:
            parent: 父容器
            total_steps: 总步骤数
        """
        self.parent = parent
        self.total_steps = total_steps
        self.current_step = 0
        self.step_titles = [f"步骤 {i+1}" for i in range(total_steps)]
        
        self.setup_ui()
    
    def setup_ui(self):
        """设置用户界面"""
        self.frame = ttk.Frame(self.parent)
        
        # 进度条标题
        self.title_label = tk.Label(
            self.frame, 
            text="", 
            font=("Arial", 12, "bold")
        )
        self.title_label.pack(pady=(0, 10))
        
        # 步骤指示器容器
        self.steps_frame = ttk.Frame(self.frame)
        self.steps_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 创建步骤指示器
        self.step_indicators = []
        self.step_lines = []
        
        for i in range(self.total_steps):
            # 创建单个步骤容器
            step_container = ttk.Frame(self.steps_frame)
            step_container.pack(side=tk.LEFT, expand=True, fill=tk.X)
            
            # 步骤圆圈和连接线的容器
            indicator_frame = ttk.Frame(step_container)
            indicator_frame.pack(pady=(0, 5))
            
            # 前置连接线（除第一个步骤外）
            if i > 0:
                line_canvas = tk.Canvas(
                    indicator_frame, 
                    width=50, 
                    height=4, 
                    highlightthickness=0,
                    bg='SystemButtonFace'
                )
                line_canvas.pack(side=tk.LEFT)
                
                line = line_canvas.create_line(0, 2, 50, 2, width=2, fill='#cccccc')
                self.step_lines.append((line_canvas, line))
            
            # 步骤圆圈
            circle_canvas = tk.Canvas(
                indicator_frame, 
                width=30, 
                height=30, 
                highlightthickness=0,
                bg='SystemButtonFace'
            )
            circle_canvas.pack(side=tk.LEFT)
            
            circle = circle_canvas.create_oval(2, 2, 28, 28, outline='#cccccc', fill='white', width=2)
            number = circle_canvas.create_text(15, 15, text=str(i+1), font=("Arial", 10, "bold"), fill='#666666')
            
            self.step_indicators.append((circle_canvas, circle, number))
            
            # 步骤标题
            step_title = tk.Label(
                step_container, 
                text=self.step_titles[i], 
                font=("Arial", 9),
                foreground='#666666'
            )
            step_title.pack()
        
        self.update_display()
    
    def set_titles(self, titles):
        """设置步骤标题
        
        Args:
            titles: 步骤标题列表
        """
        if len(titles) != self.total_steps:
            raise ValueError(f"标题数量({len(titles)})与步骤数量({self.total_steps})不匹配")
        
        self.step_titles = titles
        
        # 更新显示的标题
        for i, (step_container) in enumerate(self.steps_frame.winfo_children()):
            # 获取最后一个子控件（标题标签）
            children = step_container.winfo_children()
            if children:
                title_label = children[-1]  # 最后一个是标题标签
                if hasattr(title_label, 'config'):
                    title_label.config(text=titles[i])
    
    def set_current_step(self, step):
        """设置当前步骤
        
        Args:
            step: 当前步骤索引（0开始）
        """
        if 0 <= step < self.total_steps:
            self.current_step = step
            self.update_display()
    
    def update_display(self):
        """更新进度条显示"""
        # 更新主标题
        current_title = self.step_titles[self.current_step] if self.current_step < len(self.step_titles) else ""
        progress_text = f"{current_title} ({self.current_step + 1}/{self.total_steps})"
        self.title_label.config(text=progress_text)
        
        # 更新步骤指示器
        for i, (canvas, circle, number) in enumerate(self.step_indicators):
            if i < self.current_step:
                # 已完成步骤 - 绿色
                canvas.itemconfig(circle, outline='#4CAF50', fill='#4CAF50')
                canvas.itemconfig(number, fill='white')
            elif i == self.current_step:
                # 当前步骤 - 蓝色
                canvas.itemconfig(circle, outline='#2196F3', fill='#2196F3')
                canvas.itemconfig(number, fill='white')
            else:
                # 未完成步骤 - 灰色
                canvas.itemconfig(circle, outline='#cccccc', fill='white')
                canvas.itemconfig(number, fill='#666666')
        
        # 更新连接线
        for i, (line_canvas, line) in enumerate(self.step_lines):
            if i < self.current_step:
                # 已完成的连接线 - 绿色
                line_canvas.itemconfig(line, fill='#4CAF50')
            else:
                # 未完成的连接线 - 灰色
                line_canvas.itemconfig(line, fill='#cccccc')