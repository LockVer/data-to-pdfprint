#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件上传步骤组件
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys

# 添加项目根目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.join(current_dir, '..', '..', '..')
sys.path.insert(0, project_root)

from gui.components.base_step import BaseStep
from core.app_data import ExcelData


class FileUploadStep(BaseStep):
    """文件上传步骤"""
    
    def __init__(self, parent, app_data, title="上传Excel文件"):
        self.file_path_var = tk.StringVar()
        self.preview_text = None
        super().__init__(parent, app_data, title)
    
    def setup_content(self):
        """设置步骤具体内容"""
        # 说明文本
        info_label = tk.Label(
            self.content_frame, 
            text="请选择包含游戏数据的Excel文件（.xlsx或.xls格式）",
            font=("Arial", 11)
        )
        info_label.pack(pady=(0, 20))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(self.content_frame, text="文件选择", padding=20)
        file_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 文件路径显示
        path_frame = ttk.Frame(file_frame)
        path_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(path_frame, text="选择的文件:").pack(anchor=tk.W)
        
        path_display_frame = ttk.Frame(path_frame)
        path_display_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.path_entry = tk.Entry(
            path_display_frame, 
            textvariable=self.file_path_var,
            state='readonly',
            font=("Arial", 10)
        )
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 选择按钮
        select_btn = ttk.Button(
            path_display_frame, 
            text="浏览...", 
            command=self.select_file,
            width=10
        )
        select_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        # 文件信息预览区域
        preview_frame = ttk.LabelFrame(self.content_frame, text="文件预览", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建预览文本框
        preview_container = ttk.Frame(preview_frame)
        preview_container.pack(fill=tk.BOTH, expand=True)
        
        self.preview_text = tk.Text(
            preview_container, 
            height=12,
            wrap=tk.WORD,
            state='disabled',
            font=("Arial", 10),
            bg='#f8f8f8'
        )
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 滚动条
        preview_scroll = ttk.Scrollbar(
            preview_container, 
            orient=tk.VERTICAL, 
            command=self.preview_text.yview
        )
        preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_text.config(yscrollcommand=preview_scroll.set)
        
        # 初始显示
        self.update_preview()
    
    def select_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
        )
        
        if file_path:
            self.file_path_var.set(file_path)
            self.update_preview()
            self.clear_error()
    
    def update_preview(self):
        """更新文件预览"""
        
        file_path = self.file_path_var.get()
        
        # 清除预览内容
        self.preview_text.config(state='normal')
        self.preview_text.delete(1.0, tk.END)
        
        if not file_path:
            self.preview_text.insert(tk.END, "未选择文件")
            self.preview_text.config(state='disabled')
            return
        
        if not os.path.exists(file_path):
            self.preview_text.insert(tk.END, "文件不存在")
            self.preview_text.config(state='disabled')
            return
        
        try:
            # 确保路径正确设置
            import sys
            project_root = os.path.join(os.path.dirname(__file__), '..', '..', '..')
            src_path = os.path.join(project_root, 'src')
            if src_path not in sys.path:
                sys.path.insert(0, src_path)
                
            # 尝试读取Excel文件
            from data.excel_reader import ExcelReader
            
            reader = ExcelReader(file_path)
            excel_variables = reader.extract_template_variables()
            
            # 显示提取的信息
            info_text = f"文件信息:\n"
            info_text += f"文件名: {os.path.basename(file_path)}\n"
            info_text += f"文件大小: {self.format_file_size(os.path.getsize(file_path))}\n\n"
            
            info_text += "提取的数据:\n"
            info_text += f"客户编码 (A3): {excel_variables.get('customer_code', '未找到')}\n"
            info_text += f"主题 (B3): {excel_variables.get('theme', '未找到')}\n"
            info_text += f"开始号 (B10): {excel_variables.get('start_number', '未找到')}\n"
            info_text += f"总张数 (F3): {excel_variables.get('total_sheets', '未找到')}\n\n"
            
            # 数据验证
            missing_fields = []
            if not excel_variables.get('customer_code'):
                missing_fields.append('客户编码 (A3)')
            if not excel_variables.get('theme'):
                missing_fields.append('主题 (B3)')
            if not excel_variables.get('start_number'):
                missing_fields.append('开始号 (B10)')
            if not excel_variables.get('total_sheets'):
                missing_fields.append('总张数 (F3)')
            
            if missing_fields:
                info_text += "⚠️ 缺失字段:\n"
                for field in missing_fields:
                    info_text += f"  - {field}\n"
                info_text += "\n请检查Excel文件中相应位置的数据。"
            else:
                info_text += "✅ 所有必需字段都已找到，可以继续下一步。"
            
            self.preview_text.insert(tk.END, info_text)
            
        except Exception as e:
            error_text = f"读取Excel文件时出错:\n{str(e)}\n\n"
            error_text += "请确保:\n"
            error_text += "1. 文件是有效的Excel格式 (.xlsx 或 .xls)\n"
            error_text += "2. 文件没有被其他程序占用\n"
            error_text += "3. 文件包含必要的数据字段"
            
            self.preview_text.insert(tk.END, error_text)
        
        self.preview_text.config(state='disabled')
    
    def format_file_size(self, size_bytes):
        """格式化文件大小"""
        if size_bytes == 0:
            return "0B"
        
        size_names = ["B", "KB", "MB", "GB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
        
        return f"{size_bytes:.1f} {size_names[i]}"
    
    def validate(self) -> bool:
        """验证文件选择"""
        
        file_path = self.file_path_var.get()
        
        if not file_path:
            self.show_error("请选择Excel文件")
            return False
        
        if not os.path.exists(file_path):
            self.show_error("文件不存在")
            return False
        
        if not file_path.lower().endswith(('.xlsx', '.xls')):
            self.show_error("请选择Excel文件（.xlsx或.xls格式）")
            return False
        
        try:
            # 确保路径正确设置
            import sys
            project_root = os.path.join(os.path.dirname(__file__), '..', '..', '..')
            src_path = os.path.join(project_root, 'src')
            if src_path not in sys.path:
                sys.path.insert(0, src_path)
                
            # 尝试读取Excel文件
            from data.excel_reader import ExcelReader
            reader = ExcelReader(file_path)
            excel_variables = reader.extract_template_variables()
            
            # 检查必需字段
            required_fields = ['customer_code', 'theme', 'start_number', 'total_sheets']
            missing_fields = []
            
            for field in required_fields:
                if not excel_variables.get(field):
                    field_names = {
                        'customer_code': '客户编码 (A3)',
                        'theme': '主题 (B3)', 
                        'start_number': '开始号 (B10)',
                        'total_sheets': '总张数 (F3)'
                    }
                    missing_fields.append(field_names[field])
            
            if missing_fields:
                self.show_error(f"Excel文件缺少必需字段: {', '.join(missing_fields)}")
                return False
            
            return True
            
        except Exception as e:
            self.show_error(f"读取Excel文件失败: {str(e)}")
            return False
    
    def save_data(self):
        """保存文件数据到app_data"""
        file_path = self.file_path_var.get()
        
        try:
            # 确保路径正确设置
            import sys
            project_root = os.path.join(os.path.dirname(__file__), '..', '..', '..')
            src_path = os.path.join(project_root, 'src')
            if src_path not in sys.path:
                sys.path.insert(0, src_path)
                
            # 读取Excel数据
            from data.excel_reader import ExcelReader
            reader = ExcelReader(file_path)
            excel_variables = reader.extract_template_variables()
            
            # 保存到app_data
            self.app_data.excel_file_path = file_path
            self.app_data.excel_data = ExcelData(
                customer_code=excel_variables.get('customer_code', ''),
                theme=excel_variables.get('theme', ''),
                start_number=excel_variables.get('start_number', ''),
                total_sheets=excel_variables.get('total_sheets', 0),
                file_path=file_path
            )
            
        except Exception as e:
            raise Exception(f"保存Excel数据失败: {str(e)}")
    
    def on_show(self):
        """当步骤显示时调用"""
        super().on_show()
        
        # 如果已经有数据，恢复显示
        if self.app_data.excel_file_path:
            self.file_path_var.set(self.app_data.excel_file_path)
            self.update_preview()