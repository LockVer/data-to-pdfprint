#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF生成结果步骤组件
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import threading
import subprocess
import platform

# 添加项目根目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.join(current_dir, '..', '..', '..')
sys.path.insert(0, project_root)

from gui.components.base_step import BaseStep


class GenerateResultStep(BaseStep):
    """PDF生成结果步骤"""
    
    def __init__(self, parent, app_data, title="生成PDF文件"):
        self.is_generating = False
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="准备生成PDF文件...")
        self.result_files = []
        super().__init__(parent, app_data, title)
    
    def setup_content(self):
        """设置步骤具体内容"""
        # 生成状态区域
        status_frame = ttk.LabelFrame(self.content_frame, text="生成状态", padding=20)
        status_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 状态文本
        self.status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            font=("Arial", 11)
        )
        self.status_label.pack(pady=(0, 10))
        
        # 进度条
        self.progress_bar = ttk.Progressbar(
            status_frame,
            variable=self.progress_var,
            maximum=100,
            length=400,
            mode='determinate'
        )
        self.progress_bar.pack(pady=(0, 10))
        
        # 生成按钮
        self.generate_btn = ttk.Button(
            status_frame,
            text="开始生成",
            command=self.start_generation,
            state='normal'
        )
        self.generate_btn.pack()
        
        # 结果区域
        result_frame = ttk.LabelFrame(self.content_frame, text="生成结果", padding=15)
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        # 结果文本
        result_container = ttk.Frame(result_frame)
        result_container.pack(fill=tk.BOTH, expand=True)
        
        self.result_text = tk.Text(
            result_container,
            height=12,
            wrap=tk.WORD,
            state='disabled',
            font=("Arial", 10),
            bg='#f8f8f8'
        )
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 滚动条
        result_scroll = ttk.Scrollbar(
            result_container,
            orient=tk.VERTICAL,
            command=self.result_text.yview
        )
        result_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=result_scroll.set)
        
        # 操作按钮区域
        button_frame = ttk.Frame(result_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.open_folder_btn = ttk.Button(
            button_frame,
            text="打开输出文件夹",
            command=self.open_output_folder,
            state='disabled'
        )
        self.open_folder_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.save_as_btn = ttk.Button(
            button_frame,
            text="另存为...",
            command=self.save_files_as,
            state='disabled'
        )
        self.save_as_btn.pack(side=tk.LEFT)
    
    def start_generation(self):
        """开始生成PDF"""
        if self.is_generating:
            return
        
        # 最终验证
        if not self.app_data.is_ready_for_generation():
            errors = self.app_data.get_validation_errors()
            messagebox.showerror("配置错误", "配置不完整:\n" + "\n".join(errors))
            return
        
        self.is_generating = True
        self.app_data.is_generating = True
        
        # 更新UI状态
        self.generate_btn.config(state='disabled', text="生成中...")
        self.progress_var.set(0)
        self.status_var.set("准备开始生成...")
        
        # 清空结果
        self.result_text.config(state='normal')
        self.result_text.delete(1.0, tk.END)
        self.result_text.config(state='disabled')
        
        # 在新线程中执行生成
        threading.Thread(target=self._generate_pdfs, daemon=True).start()
    
    def _generate_pdfs(self):
        """在后台线程中生成PDF"""
        try:
            # 更新状态
            self.root.after(0, lambda: self.update_status("正在初始化...", 10))
            
            # 准备输出目录
            folder_name, box_name, case_name = self.app_data.get_output_file_names()
            output_dir = os.path.join(os.getcwd(), folder_name)
            
            # 创建输出目录
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            self.root.after(0, lambda: self.update_status("正在生成盒标PDF...", 30))
            
            # 生成盒标
            box_pdf_path = os.path.join(output_dir, box_name)
            self._generate_box_labels(box_pdf_path)
            
            self.root.after(0, lambda: self.update_status("正在生成箱标PDF...", 70))
            
            # 生成箱标  
            case_pdf_path = os.path.join(output_dir, case_name)
            self._generate_case_labels(case_pdf_path)
            
            self.root.after(0, lambda: self.update_status("生成完成!", 100))
            
            # 保存结果
            self.result_files = [box_pdf_path, case_pdf_path]
            self.app_data.generated_files = self.result_files
            self.app_data.output_directory = output_dir
            
            # 更新结果显示
            self.root.after(0, self._on_generation_success)
            
        except Exception as e:
            error_msg = str(e)
            self.app_data.generation_error = error_msg
            self.root.after(0, lambda: self._on_generation_error(error_msg))
    
    def _generate_box_labels(self, output_path):
        """生成盒标PDF"""
        try:
            # 确保路径正确设置
            import sys
            project_root = os.path.join(os.path.dirname(__file__), '..', '..', '..')
            src_path = os.path.join(project_root, 'src')
            if src_path not in sys.path:
                sys.path.insert(0, src_path)
            
            from render_pdf import generate_pdf_from_excel
            
            # 准备参数
            additional_inputs = {
                'sheets_per_box': self.app_data.package_params.sheets_per_box,
                'boxes_per_small_case': self.app_data.package_params.boxes_per_small_case,
                'small_cases_per_large_case': self.app_data.package_params.small_cases_per_large_case
            }
            
            template_mode = "game_info" if self.app_data.label_template == "game" else "two_level"
            
            # 生成PDF
            generate_pdf_from_excel(
                self.app_data.excel_file_path,
                output_path,
                additional_inputs,
                template_mode
            )
            
        except Exception as e:
            raise Exception(f"生成盒标失败: {str(e)}")
    
    def _generate_case_labels(self, output_path):
        """生成箱标PDF"""
        try:
            # 确保路径正确设置
            import sys
            project_root = os.path.join(os.path.dirname(__file__), '..', '..', '..')
            src_path = os.path.join(project_root, 'src')
            if src_path not in sys.path:
                sys.path.insert(0, src_path)
                
            from render_case_pdf import generate_case_pdf_from_excel
            
            # 准备参数
            additional_inputs = {
                'sheets_per_box': self.app_data.package_params.sheets_per_box,
                'boxes_per_small_case': self.app_data.package_params.boxes_per_small_case,
                'small_cases_per_large_case': self.app_data.package_params.small_cases_per_large_case
            }
            
            # 根据包装模式选择箱标模板
            template_mode_map = {
                'regular': 'regular_single',
                'separate': 'multi_separate',
                'set': 'multi_set'
            }
            template_mode = template_mode_map.get(self.app_data.package_mode, 'regular_single')
            
            # 生成PDF
            generate_case_pdf_from_excel(
                self.app_data.excel_file_path,
                output_path,
                additional_inputs,
                template_mode
            )
            
        except Exception as e:
            raise Exception(f"生成箱标失败: {str(e)}")
    
    def update_status(self, status, progress):
        """更新状态和进度"""
        self.status_var.set(status)
        self.progress_var.set(progress)
    
    def _on_generation_success(self):
        """生成成功处理"""
        self.is_generating = False
        self.app_data.is_generating = False
        
        # 更新UI状态
        self.generate_btn.config(state='normal', text="重新生成")
        self.open_folder_btn.config(state='normal')
        self.save_as_btn.config(state='normal')
        
        # 显示结果
        result_text = f"✅ PDF文件生成成功！\n\n"
        result_text += f"生成位置: {self.app_data.output_directory}\n\n"
        result_text += f"生成的文件:\n"
        
        for file_path in self.result_files:
            file_name = os.path.basename(file_path)
            file_size = self.format_file_size(os.path.getsize(file_path))
            result_text += f"  📄 {file_name} ({file_size})\n"
        
        result_text += f"\n生成统计:\n"
        
        # 添加生成统计信息
        if self.app_data.excel_data and self.app_data.package_params:
            import math
            total_sheets = self.app_data.excel_data.total_sheets
            sheets_per_box = self.app_data.package_params.sheets_per_box
            total_boxes = math.ceil(total_sheets / sheets_per_box)
            
            result_text += f"  • 总卡片数: {total_sheets:,} 张\n"
            result_text += f"  • 生成盒数: {total_boxes} 盒\n"
            result_text += f"  • 包装模式: {self.app_data.get_package_mode_display_name()}\n"
            result_text += f"  • 标签模板: {self.app_data.get_label_template_display_name()}\n"
        
        result_text += f"\n💡 提示:\n"
        result_text += f"  • 点击「打开输出文件夹」查看生成的文件\n"
        result_text += f"  • 点击「另存为」将文件保存到其他位置\n"
        result_text += f"  • PDF格式为PDF/X，适合CMYK打印\n"
        
        self.result_text.config(state='normal')
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, result_text)
        self.result_text.config(state='disabled')
    
    def _on_generation_error(self, error_msg):
        """生成失败处理"""
        self.is_generating = False
        self.app_data.is_generating = False
        
        # 更新UI状态
        self.generate_btn.config(state='normal', text="重新生成")
        self.status_var.set("生成失败")
        
        # 显示错误信息
        result_text = f"❌ PDF生成失败\n\n"
        result_text += f"错误信息:\n{error_msg}\n\n"
        result_text += f"可能的解决方案:\n"
        result_text += f"  • 检查Excel文件是否正确\n"
        result_text += f"  • 确保输出目录有写入权限\n"
        result_text += f"  • 检查网络连接和系统资源\n"
        result_text += f"  • 重新选择参数配置\n"
        
        self.result_text.config(state='normal')
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, result_text)
        self.result_text.config(state='disabled')
    
    def open_output_folder(self):
        """打开输出文件夹"""
        if not self.app_data.output_directory or not os.path.exists(self.app_data.output_directory):
            messagebox.showerror("错误", "输出文件夹不存在")
            return
        
        try:
            # 跨平台打开文件夹
            if platform.system() == "Windows":
                os.startfile(self.app_data.output_directory)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", self.app_data.output_directory])
            else:  # Linux
                subprocess.run(["xdg-open", self.app_data.output_directory])
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹: {str(e)}")
    
    def save_files_as(self):
        """另存为文件"""
        if not self.result_files:
            messagebox.showwarning("提示", "没有可保存的文件")
            return
        
        try:
            # 选择保存目录
            save_dir = filedialog.askdirectory(title="选择保存位置")
            if not save_dir:
                return
            
            # 复制文件
            import shutil
            
            saved_files = []
            for src_file in self.result_files:
                if os.path.exists(src_file):
                    file_name = os.path.basename(src_file)
                    dst_file = os.path.join(save_dir, file_name)
                    shutil.copy2(src_file, dst_file)
                    saved_files.append(dst_file)
            
            messagebox.showinfo(
                "保存成功", 
                f"已保存 {len(saved_files)} 个文件到:\n{save_dir}"
            )
            
        except Exception as e:
            messagebox.showerror("保存失败", f"保存文件时出错: {str(e)}")
    
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
        """验证生成步骤"""
        return self.app_data.is_ready_for_generation()
    
    def save_data(self):
        """保存数据（生成步骤无需保存）"""
        pass
    
    def on_show(self):
        """当步骤显示时调用"""
        super().on_show()
        
        # 更新生成按钮状态
        if self.app_data.is_ready_for_generation():
            self.generate_btn.config(state='normal')
            self.status_var.set("准备就绪，点击开始生成PDF文件")
        else:
            self.generate_btn.config(state='disabled')
            errors = self.app_data.get_validation_errors()
            self.status_var.set(f"配置不完整: {'; '.join(errors)}")
    
    @property
    def root(self):
        """获取根窗口"""
        widget = self.frame
        while widget.master:
            widget = widget.master
        return widget