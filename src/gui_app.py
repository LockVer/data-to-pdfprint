"""
GUI应用程序
支持选择Excel文件进行处理
支持Windows和macOS
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from pathlib import Path
import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from src.pdf.generator import PDFGenerator


class DataToPDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data to PDF Print - Excel转PDF工具")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # 设置窗口居中
        self.center_window()
        
        self.setup_ui()
        self.setup_file_selection()
        
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 标题
        title_label = ttk.Label(main_frame, text="Excel数据到PDF转换工具", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # 文件选择区域
        self.select_frame = tk.Frame(main_frame, bg="#f0f0f0", 
                                  relief="ridge", bd=2, height=120)
        self.select_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        self.select_frame.grid_propagate(False)
        
        # 文件选择提示
        self.select_label = tk.Label(self.select_frame, 
                             text="📁 点击此区域选择Excel文件\n\n支持 .xlsx 和 .xls 格式",
                             bg="#f0f0f0", font=("Arial", 11), fg="#666666", cursor="hand2")
        self.select_label.place(relx=0.5, rely=0.5, anchor="center")
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, pady=(0, 20))
        
        # 选择文件按钮
        select_btn = ttk.Button(button_frame, text="📂 选择Excel文件", 
                               command=self.select_file)
        select_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 生成PDF按钮
        self.generate_btn = ttk.Button(button_frame, text="🔄 生成PDF", 
                                      command=self.generate_pdf, state="disabled")
        self.generate_btn.pack(side=tk.LEFT)
        
        # 文件信息显示
        info_frame = ttk.LabelFrame(main_frame, text="文件信息", padding="10")
        info_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        
        self.info_text = tk.Text(info_frame, height=10, width=70, font=("Consolas", 10))
        self.info_text.pack(fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(info_frame, orient="vertical", command=self.info_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.info_text.configure(yscrollcommand=scrollbar.set)
        
        # 状态栏
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        self.status_var = tk.StringVar()
        self.status_var.set("📋 准备就绪 - 请选择Excel文件")
        status_label = ttk.Label(status_frame, textvariable=self.status_var)
        status_label.pack(side=tk.LEFT)
        
        # 配置网格权重
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(3, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        
        self.current_file = None
        self.current_data = None
        
    def center_window(self):
        """窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def setup_file_selection(self):
        """设置文件选择功能"""
        # 点击区域打开文件选择
        self.select_frame.bind("<Button-1>", self.on_click_select)
        self.select_label.bind("<Button-1>", self.on_click_select)
        self.root.bind("<Control-o>", lambda e: self.select_file())  # Ctrl+O快捷键
        
    def on_click_select(self, event):
        """点击选择区域打开文件选择"""
        self.select_file()
    
    def select_file(self):
        """选择文件对话框"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.process_file(file_path)
    
    def process_file(self, file_path):
        """处理Excel文件"""
        try:
            self.status_var.set("🔄 正在处理文件...")
            self.info_text.delete(1.0, tk.END)
            self.info_text.insert(tk.END, "正在读取Excel文件...\n")
            self.root.update()
            
            # 检查文件格式
            if not file_path.lower().endswith(('.xlsx', '.xls')):
                messagebox.showerror("格式错误", "请选择Excel文件(.xlsx或.xls)")
                self.status_var.set("❌ 文件格式错误")
                return
            
            # 读取Excel文件
            df = pd.read_excel(file_path, header=None)
            
            # 提取数据
            total_count = df.iloc[3,5]  # 总张数
            
            self.current_data = {
                '客户编码': str(df.iloc[3,0]),
                '主题': str(df.iloc[3,1]), 
                '排列要求': str(df.iloc[3,2]),
                '订单数量': str(df.iloc[3,3]),
                '张/盒': str(df.iloc[3,4]),
                '总张数': str(total_count)
            }
            
            # 显示提取的信息
            info_text = f"文件: {Path(file_path).name}\n"
            info_text += f"文件大小: {Path(file_path).stat().st_size} 字节\n\n"
            info_text += "提取的数据:\n"
            info_text += "-" * 40 + "\n"
            
            for key, value in self.current_data.items():
                info_text += f"{key}: {value}\n"
            
            self.info_text.insert(tk.END, info_text)
            
            self.current_file = file_path
            self.generate_btn.config(state="normal")
            self.status_var.set(f"✅ 文件处理完成 - 总张数: {total_count}")
            
            # 更新选择区域显示
            self.select_label.config(text=f"✅ 已选择文件: {Path(file_path).name}\n总张数: {total_count}\n\n点击生成PDF按钮继续", 
                                  fg="green")
            
        except Exception as e:
            error_msg = f"处理文件失败: {str(e)}"
            messagebox.showerror("处理错误", error_msg)
            self.status_var.set("❌ 处理失败")
            self.info_text.insert(tk.END, f"\n错误: {error_msg}\n")
    
    def generate_pdf(self):
        """生成PDF"""
        if not self.current_data:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
        
        try:
            self.status_var.set("🔄 正在生成PDF...")
            self.info_text.insert(tk.END, "\n开始生成PDF...\n")
            self.root.update()
            
            # 选择输出位置 - 支持自定义文件夹
            default_name = f"label_{self.current_data.get('客户编码', 'output')}.pdf"
            output_path = filedialog.asksaveasfilename(
                title="选择保存位置和文件名",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
                initialfile=default_name
            )
            
            if output_path:
                # 生成PDF
                generator = PDFGenerator()
                generator.create_label_pdf(self.current_data, output_path)
                
                self.status_var.set(f"✅ PDF生成成功!")
                self.info_text.insert(tk.END, f"PDF已保存到: {output_path}\n")
                
                # 询问是否打开文件
                if messagebox.askyesno("生成成功", f"PDF已生成!\n\n保存位置: {output_path}\n\n是否打开文件？"):
                    import subprocess
                    import platform
                    
                    if platform.system() == "Darwin":  # macOS
                        subprocess.run(["open", output_path])
                    elif platform.system() == "Windows":  # Windows
                        os.startfile(output_path)
                        
            else:
                self.status_var.set("📋 PDF生成已取消")
                
        except Exception as e:
            error_msg = f"生成PDF失败: {str(e)}"
            messagebox.showerror("生成错误", error_msg)
            self.status_var.set("❌ PDF生成失败")
            self.info_text.insert(tk.END, f"\n错误: {error_msg}\n")


def main():
    """启动GUI应用"""
    root = tk.Tk()
    app = DataToPDFApp(root)
    
    # 检查命令行参数，支持文件关联
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        if file_path.lower().endswith(('.xlsx', '.xls')):
            # 延迟处理文件，等GUI完全加载
            root.after(500, lambda: app.process_file(file_path))
    
    root.mainloop()


if __name__ == "__main__":
    main()