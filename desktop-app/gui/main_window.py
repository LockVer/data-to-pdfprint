#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF标签生成器 - 主窗口
"""

import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys

# 添加项目根目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.join(current_dir, '..', '..')
sys.path.insert(0, project_root)

from gui.components.progress_bar import ProgressBar
from gui.components.file_upload import FileUploadStep
from gui.components.mode_selector import ModeSelectorStep
from gui.components.param_input import ParamInputStep
from gui.components.template_selector import TemplateSelectorStep
from gui.components.generate_result import GenerateResultStep
from core.app_data import AppData


class PDFLabelGeneratorApp:
    """PDF标签生成器主应用程序"""
    
    def __init__(self):
        """初始化应用程序"""
        self.root = tk.Tk()
        self.current_step = 0
        self.total_steps = 5
        self.app_data = AppData()
        
        # 步骤列表
        self.steps = []
        
        self.setup_window()
        self.setup_ui()
        self.create_steps()
        self.show_current_step()
    
    def setup_window(self):
        """设置主窗口"""
        self.root.title("PDF标签生成器 v1.0")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)
        
        # 设置窗口居中
        self.center_window()
        
        # 设置窗口图标（如果存在）
        try:
            icon_path = os.path.join(os.path.dirname(__file__), '..', 'assets', 'icons', 'app.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            pass
    
    def center_window(self):
        """使窗口在屏幕中央显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)
        
        # 标题
        title_label = tk.Label(
            self.main_frame, 
            text="PDF标签生成器", 
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # 进度条
        self.progress_bar = ProgressBar(self.main_frame, self.total_steps)
        self.progress_bar.frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # 内容区域
        self.content_frame = ttk.Frame(self.main_frame)
        self.content_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        self.content_frame.columnconfigure(0, weight=1)
        self.content_frame.rowconfigure(0, weight=1)
        
        # 按钮区域
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.grid(row=3, column=0, sticky=(tk.W, tk.E))
        
        # 按钮
        self.prev_button = ttk.Button(
            self.button_frame, 
            text="上一步", 
            command=self.prev_step,
            state='disabled'
        )
        self.prev_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.next_button = ttk.Button(
            self.button_frame, 
            text="下一步", 
            command=self.next_step
        )
        self.next_button.pack(side=tk.RIGHT)
    
    def create_steps(self):
        """创建所有步骤组件"""
        step_classes = [
            FileUploadStep,
            ModeSelectorStep,
            ParamInputStep,
            TemplateSelectorStep,
            GenerateResultStep
        ]
        
        step_titles = [
            "上传Excel文件",
            "选择包装模式", 
            "输入包装参数",
            "选择标签模板",
            "生成PDF文件"
        ]
        
        for i, (step_class, title) in enumerate(zip(step_classes, step_titles)):
            step = step_class(self.content_frame, self.app_data, title)
            step.frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            step.frame.grid_remove()  # 初始隐藏
            self.steps.append(step)
        
        # 更新进度条标题
        self.progress_bar.set_titles(step_titles)
    
    def show_current_step(self):
        """显示当前步骤"""
        # 隐藏所有步骤
        for step in self.steps:
            step.frame.grid_remove()
        
        # 显示当前步骤
        if 0 <= self.current_step < len(self.steps):
            self.steps[self.current_step].frame.grid()
            self.steps[self.current_step].on_show()
        
        # 更新进度条
        self.progress_bar.set_current_step(self.current_step)
        
        # 更新按钮状态
        self.update_button_states()
    
    def update_button_states(self):
        """更新按钮状态"""
        # 上一步按钮
        if self.current_step == 0:
            self.prev_button.config(state='disabled')
        else:
            self.prev_button.config(state='normal')
        
        # 下一步/生成按钮
        if self.current_step == len(self.steps) - 1:
            self.next_button.config(text="生成PDF")
        else:
            self.next_button.config(text="下一步")
    
    def prev_step(self):
        """上一步"""
        if self.current_step > 0:
            self.current_step -= 1
            self.show_current_step()
    
    def next_step(self):
        """下一步"""
        # 验证当前步骤
        current_step_obj = self.steps[self.current_step]
        if not current_step_obj.validate():
            return
        
        # 保存当前步骤数据
        current_step_obj.save_data()
        
        if self.current_step < len(self.steps) - 1:
            self.current_step += 1
            self.show_current_step()
        else:
            # 最后一步，开始生成PDF
            self.generate_pdf()
    
    def generate_pdf(self):
        """生成PDF文件"""
        try:
            # 获取最后一步的组件并开始生成
            generate_step = self.steps[-1]
            generate_step.start_generation()
        except Exception as e:
            messagebox.showerror("生成错误", f"生成PDF时发生错误：{str(e)}")
    
    def run(self):
        """运行应用程序"""
        self.root.mainloop()


if __name__ == "__main__":
    app = PDFLabelGeneratorApp()
    app.run()