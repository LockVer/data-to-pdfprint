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
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

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
        title_label = ttk.Label(
            main_frame, text="Excel数据到PDF转换工具", font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 20))

        # 文件选择区域
        self.select_frame = tk.Frame(
            main_frame, bg="#f0f0f0", relief="ridge", bd=2, height=120
        )
        self.select_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        self.select_frame.grid_propagate(False)

        # 文件选择提示
        self.select_label = tk.Label(
            self.select_frame,
            text="📁 点击此区域选择Excel文件\n\n支持 .xlsx 和 .xls 格式",
            bg="#f0f0f0",
            font=("Arial", 11),
            fg="#666666",
            cursor="hand2",
        )
        self.select_label.place(relx=0.5, rely=0.5, anchor="center")

        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, pady=(0, 20))

        # 选择文件按钮
        select_btn = ttk.Button(
            button_frame, text="📂 选择Excel文件", command=self.select_file
        )
        select_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 生成PDF按钮
        self.generate_btn = ttk.Button(
            button_frame,
            text="🔄 选择模板并生成PDF",
            command=self.start_generation_workflow,
            state="disabled",
        )
        self.generate_btn.pack(side=tk.LEFT)

        # 文件信息显示
        info_frame = ttk.LabelFrame(main_frame, text="文件信息", padding="10")
        info_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))

        self.info_text = tk.Text(info_frame, height=10, width=70, font=("Consolas", 10))
        self.info_text.pack(fill=tk.BOTH, expand=True)

        # 滚动条
        scrollbar = ttk.Scrollbar(
            info_frame, orient="vertical", command=self.info_text.yview
        )
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
        self.packaging_params = None

    def _extract_total_count_by_keyword(self, df):
        """通过关键字搜索提取总张数"""
        try:
            # 搜索包含"总张数"的单元格
            for i in range(df.shape[0]):
                for j in range(df.shape[1]):
                    cell_value = df.iloc[i, j]
                    if pd.notna(cell_value) and "总张数" in str(cell_value):
                        print(f"✅ 找到总张数关键字: 位置({i+1},{j+1}) = '{cell_value}'")
                        
                        # 尝试从下方单元格获取数值
                        if i + 1 < df.shape[0]:
                            total_value = df.iloc[i + 1, j]
                            if pd.notna(total_value):
                                print(f"✅ 从下方提取总张数: {total_value}")
                                return int(float(total_value))
                        
                        # 如果下方没有数据，尝试同行右侧
                        if j + 1 < df.shape[1]:
                            total_value = df.iloc[i, j + 1]
                            if pd.notna(total_value):
                                print(f"✅ 从右侧提取总张数: {total_value}")
                                return int(float(total_value))
                        
                        # 最后尝试同行后几列
                        for k in range(j + 1, min(j + 5, df.shape[1])):
                            total_value = df.iloc[i, k]
                            if pd.notna(total_value) and str(total_value).replace('.', '').replace('-', '').isdigit():
                                print(f"✅ 从右侧第{k-j}列提取总张数: {total_value}")
                                return int(float(total_value))
            
            # 如果没找到关键字，使用默认位置
            print("⚠️ 未找到总张数关键字，使用默认位置(4,6)")
            default_value = df.iloc[3, 5]
            if pd.notna(default_value):
                return int(float(default_value))
            else:
                print("❌ 默认位置也无数据，返回0")
                return 0
                
        except Exception as e:
            print(f"❌ 提取总张数失败: {e}")
            return 0

    def center_window(self):
        """窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

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
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
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
            if not file_path.lower().endswith((".xlsx", ".xls")):
                messagebox.showerror("格式错误", "请选择Excel文件(.xlsx或.xls)")
                self.status_var.set("❌ 文件格式错误")
                return

            # 读取Excel文件
            df = pd.read_excel(file_path, header=None)

            # 使用关键字搜索提取总张数
            total_count = self._extract_total_count_by_keyword(df)

            self.current_data = {
                "客户编码": str(df.iloc[3, 0]),
                "主题": str(df.iloc[3, 1]),
                "排列要求": str(df.iloc[3, 2]),
                "订单数量": str(df.iloc[3, 3]),
                "张/盒": str(df.iloc[3, 4]),
                "总张数": str(total_count),
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
            display_text = (
                f"✅ 已选择文件: {Path(file_path).name}\n总张数: {total_count}"
                f"\n\n点击生成多级标签PDF按钮继续"
            )
            self.select_label.config(text=display_text, fg="green")

        except Exception as e:
            error_msg = f"处理文件失败: {str(e)}"
            messagebox.showerror("处理错误", error_msg)
            self.status_var.set("❌ 处理失败")
            self.info_text.insert(tk.END, f"\n错误: {error_msg}\n")

    def show_parameters_dialog(self):
        """显示参数设置对话框"""
        if not self.current_data:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return

        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("包装参数设置")
        dialog.geometry("500x450")
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (dialog.winfo_screenheight() // 2) - (450 // 2)
        dialog.geometry(f"500x450+{x}+{y}")

        # 创建滚动框架
        canvas = tk.Canvas(dialog)
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 绑定鼠标滚轮
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 主框架在可滚动区域内
        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_label = ttk.Label(
            main_frame, text="设置包装参数", font=("Arial", 14, "bold")
        )
        title_label.pack(pady=(0, 20))

        # 参数输入框架
        params_frame = ttk.LabelFrame(main_frame, text="包装参数", padding="15")
        params_frame.pack(fill=tk.X, pady=(0, 20))

        # 张/盒输入
        ttk.Label(params_frame, text="张/盒:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        self.pieces_per_box_var = tk.StringVar(value="2850")
        pieces_per_box_entry = ttk.Entry(
            params_frame, textvariable=self.pieces_per_box_var, width=15
        )
        pieces_per_box_entry.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=5)

        # 盒/小箱输入
        ttk.Label(params_frame, text="盒/小箱:").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        self.boxes_per_small_box_var = tk.StringVar(value="1")
        boxes_per_small_box_entry = ttk.Entry(
            params_frame, textvariable=self.boxes_per_small_box_var, width=15
        )
        boxes_per_small_box_entry.grid(
            row=1, column=1, sticky=tk.W, padx=(10, 0), pady=5
        )

        # 小箱/大箱输入
        ttk.Label(params_frame, text="小箱/大箱:").grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        self.small_boxes_per_large_box_var = tk.StringVar(value="2")
        small_boxes_entry = ttk.Entry(
            params_frame, textvariable=self.small_boxes_per_large_box_var, width=15
        )
        small_boxes_entry.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=5)

        # 外观选择框架
        appearance_frame = ttk.LabelFrame(main_frame, text="盒标外观选择", padding="15")
        appearance_frame.pack(fill=tk.X, pady=(0, 20))

        self.appearance_var = tk.StringVar(value="外观一")
        appearance_radio1 = ttk.Radiobutton(
            appearance_frame,
            text="外观一",
            variable=self.appearance_var,
            value="外观一",
        )
        appearance_radio1.pack(side=tk.LEFT, padx=(0, 20))

        appearance_radio2 = ttk.Radiobutton(
            appearance_frame,
            text="外观二",
            variable=self.appearance_var,
            value="外观二",
        )
        appearance_radio2.pack(side=tk.LEFT)

        # 当前数据显示
        info_frame = ttk.LabelFrame(main_frame, text="当前数据", padding="15")
        info_frame.pack(fill=tk.X, pady=(0, 20))

        info_text = f"客户编码: {self.current_data['客户编码']}\n"
        info_text += f"主题: {self.current_data['主题']}\n"
        info_text += f"总张数: {self.current_data['总张数']}"

        info_label = ttk.Label(info_frame, text=info_text, font=("Consolas", 10))
        info_label.pack(anchor=tk.W)

        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(10, 0))

        # 确认按钮
        confirm_btn = ttk.Button(
            button_frame,
            text="确认生成",
            command=lambda: self.confirm_parameters(dialog),
        )
        confirm_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 取消按钮
        cancel_btn = ttk.Button(button_frame, text="取消", command=dialog.destroy)
        cancel_btn.pack(side=tk.LEFT)

        # 设置焦点
        pieces_per_box_entry.focus()

    def show_fenhe_parameters_dialog(self):
        """显示分盒模板的参数设置对话框（无外观选择）"""
        if not self.current_data:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return

        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("分盒模板 - 包装参数设置")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)
        dialog.geometry(f"500x400+{x}+{y}")

        # 创建滚动框架
        canvas = tk.Canvas(dialog)
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 绑定鼠标滚轮
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 主框架在可滚动区域内
        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_label = ttk.Label(
            main_frame, text="分盒模板参数设置", font=("Arial", 14, "bold")
        )
        title_label.pack(pady=(0, 20))

        # 参数输入框架
        params_frame = ttk.LabelFrame(main_frame, text="包装参数", padding="15")
        params_frame.pack(fill=tk.X, pady=(0, 20))

        # 张/盒输入
        ttk.Label(params_frame, text="张/盒:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        self.pieces_per_box_var = tk.StringVar(value="2850")
        pieces_per_box_entry = ttk.Entry(
            params_frame, textvariable=self.pieces_per_box_var, width=15
        )
        pieces_per_box_entry.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=5)

        # 盒/小箱输入
        ttk.Label(params_frame, text="盒/小箱:").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        self.boxes_per_small_box_var = tk.StringVar(value="1")
        boxes_per_small_box_entry = ttk.Entry(
            params_frame, textvariable=self.boxes_per_small_box_var, width=15
        )
        boxes_per_small_box_entry.grid(
            row=1, column=1, sticky=tk.W, padx=(10, 0), pady=5
        )

        # 小箱/大箱输入
        ttk.Label(params_frame, text="小箱/大箱:").grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        self.small_boxes_per_large_box_var = tk.StringVar(value="2")
        small_boxes_entry = ttk.Entry(
            params_frame, textvariable=self.small_boxes_per_large_box_var, width=15
        )
        small_boxes_entry.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=5)

        # 提示信息框架
        info_frame = ttk.LabelFrame(main_frame, text="分盒模板说明", padding="15")
        info_frame.pack(fill=tk.X, pady=(0, 20))

        info_text = "分盒模板使用特殊的序列号生成规则：\n"
        info_text += "• 从小箱/大箱参数控制副号满几进一\n"
        info_text += "• 序列号格式：前缀+数字-后缀\n"
        info_text += "• 示例：MOP01001-01, MOP01001-02, MOP01002-01...\n"
        info_text += "• 说明：小箱标序列号只取决于小箱/大箱参数，盒/小箱参数建议设为1"

        info_label = ttk.Label(info_frame, text=info_text, font=("Consolas", 9))
        info_label.pack(anchor=tk.W)

        # 当前数据显示
        data_frame = ttk.LabelFrame(main_frame, text="当前数据", padding="15")
        data_frame.pack(fill=tk.X, pady=(0, 20))

        data_text = f"客户编码: {self.current_data['客户编码']}\n"
        data_text += f"主题: {self.current_data['主题']}\n"
        data_text += f"总张数: {self.current_data['总张数']}"

        data_label = ttk.Label(data_frame, text=data_text, font=("Consolas", 10))
        data_label.pack(anchor=tk.W)

        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(10, 0))

        # 确认按钮
        confirm_btn = ttk.Button(
            button_frame,
            text="确认生成",
            command=lambda: self.confirm_fenhe_parameters(dialog),
        )
        confirm_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 取消按钮
        cancel_btn = ttk.Button(button_frame, text="取消", command=dialog.destroy)
        cancel_btn.pack(side=tk.LEFT)

        # 设置焦点
        pieces_per_box_entry.focus()

    def confirm_fenhe_parameters(self, dialog):
        """确认分盒模板参数并生成PDF"""
        try:
            # 验证三个参数
            pieces_per_box = int(self.pieces_per_box_var.get())
            boxes_per_small_box = int(self.boxes_per_small_box_var.get())
            small_boxes_per_large_box = int(self.small_boxes_per_large_box_var.get())

            if (
                pieces_per_box <= 0
                or boxes_per_small_box <= 0
                or small_boxes_per_large_box <= 0
            ):
                messagebox.showerror("参数错误", "所有参数必须为正整数")
                return

            # 分盒模板不需要外观选择，使用默认外观一
            self.packaging_params = {
                "张/盒": pieces_per_box,
                "盒/小箱": boxes_per_small_box,
                "小箱/大箱": small_boxes_per_large_box,
                "选择外观": "外观一",  # 分盒模板固定使用外观一
            }

            dialog.destroy()
            self.generate_multi_level_pdfs()

        except ValueError:
            messagebox.showerror("参数错误", "请输入有效的数字")

    def show_taohebox_parameters_dialog(self):
        """显示套盒模板的参数设置对话框（无外观选择）"""
        if not self.current_data:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return

        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("套盒模板 - 包装参数设置")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)
        dialog.geometry(f"500x400+{x}+{y}")

        # 创建滚动框架
        canvas = tk.Canvas(dialog)
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 绑定鼠标滚轮
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 主框架在可滚动区域内
        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_label = ttk.Label(
            main_frame, text="套盒模板参数设置", font=("Arial", 14, "bold")
        )
        title_label.pack(pady=(0, 20))

        # 参数输入框架
        params_frame = ttk.LabelFrame(main_frame, text="包装参数", padding="15")
        params_frame.pack(fill=tk.X, pady=(0, 20))

        # 张/盒输入
        ttk.Label(params_frame, text="张/盒:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        self.pieces_per_box_var = tk.StringVar(value="2850")
        pieces_per_box_entry = ttk.Entry(
            params_frame, textvariable=self.pieces_per_box_var, width=15
        )
        pieces_per_box_entry.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=5)

        # 盒/小箱输入
        ttk.Label(params_frame, text="盒/小箱:").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        self.boxes_per_small_box_var = tk.StringVar(value="1")
        boxes_per_small_box_entry = ttk.Entry(
            params_frame, textvariable=self.boxes_per_small_box_var, width=15
        )
        boxes_per_small_box_entry.grid(
            row=1, column=1, sticky=tk.W, padx=(10, 0), pady=5
        )

        # 小箱/大箱输入
        ttk.Label(params_frame, text="小箱/大箱:").grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        self.small_boxes_per_large_box_var = tk.StringVar(value="2")
        small_boxes_entry = ttk.Entry(
            params_frame, textvariable=self.small_boxes_per_large_box_var, width=15
        )
        small_boxes_entry.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=5)

        # 提示信息框架
        info_frame = ttk.LabelFrame(main_frame, text="套盒模板说明", padding="15")
        info_frame.pack(fill=tk.X, pady=(0, 20))

        info_text = "套盒模板使用Excel文件中的开始号和结束号：\\n"
        info_text += "• 第二个参数(盒/小箱)用于控制结束号范围\\n"
        info_text += "• 序列号格式基于Excel文件中的开始号和结束号\\n"
        info_text += "• 示例：开始号JAW01001-01，结束号JAW01001-06\\n"
        info_text += "• 说明：套盒模板无外观选择，使用固定外观"

        info_label = ttk.Label(info_frame, text=info_text, font=("Consolas", 9))
        info_label.pack(anchor=tk.W)

        # 当前数据显示
        data_frame = ttk.LabelFrame(main_frame, text="当前数据", padding="15")
        data_frame.pack(fill=tk.X, pady=(0, 20))

        data_text = f"客户编码: {self.current_data['客户编码']}\\n"
        data_text += f"主题: {self.current_data['主题']}\\n"
        data_text += f"总张数: {self.current_data['总张数']}"

        data_label = ttk.Label(data_frame, text=data_text, font=("Consolas", 10))
        data_label.pack(anchor=tk.W)

        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(10, 0))

        # 确认按钮
        confirm_btn = ttk.Button(
            button_frame,
            text="确认生成",
            command=lambda: self.confirm_taohebox_parameters(dialog),
        )
        confirm_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 取消按钮
        cancel_btn = ttk.Button(button_frame, text="取消", command=dialog.destroy)
        cancel_btn.pack(side=tk.LEFT)

        # 设置焦点
        pieces_per_box_entry.focus()

    def confirm_taohebox_parameters(self, dialog):
        """确认套盒模板参数并生成PDF"""
        try:
            # 验证三个参数
            pieces_per_box = int(self.pieces_per_box_var.get())
            boxes_per_small_box = int(self.boxes_per_small_box_var.get())
            small_boxes_per_large_box = int(self.small_boxes_per_large_box_var.get())

            if (
                pieces_per_box <= 0
                or boxes_per_small_box <= 0
                or small_boxes_per_large_box <= 0
            ):
                messagebox.showerror("参数错误", "所有参数必须为正整数")
                return

            # 套盒模板不需要外观选择，使用默认外观
            self.packaging_params = {
                "张/盒": pieces_per_box,
                "盒/小箱": boxes_per_small_box,
                "小箱/大箱": small_boxes_per_large_box,
                "选择外观": "外观一",  # 套盒模板固定使用外观一，但实际不使用
            }

            dialog.destroy()
            self.generate_multi_level_pdfs()

        except ValueError:
            messagebox.showerror("参数错误", "请输入有效的数字")

    def confirm_parameters(self, dialog):
        """确认参数并生成PDF"""
        try:
            # 验证三个参数
            pieces_per_box = int(self.pieces_per_box_var.get())
            boxes_per_small_box = int(self.boxes_per_small_box_var.get())
            small_boxes_per_large_box = int(self.small_boxes_per_large_box_var.get())

            if (
                pieces_per_box <= 0
                or boxes_per_small_box <= 0
                or small_boxes_per_large_box <= 0
            ):
                messagebox.showerror("参数错误", "所有参数必须为正整数")
                return

            # 获取选择的外观
            selected_appearance = self.appearance_var.get()

            self.packaging_params = {
                "张/盒": pieces_per_box,
                "盒/小箱": boxes_per_small_box,
                "小箱/大箱": small_boxes_per_large_box,
                "选择外观": selected_appearance,
            }

            dialog.destroy()
            self.generate_multi_level_pdfs()

        except ValueError:
            messagebox.showerror("参数错误", "请输入有效的数字")

    def show_template_selection_dialog(self):
        """显示模板选择对话框"""
        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("选择标签模板")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (dialog.winfo_screenheight() // 2) - (300 // 2)
        dialog.geometry(f"400x300+{x}+{y}")

        # 主框架
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_label = ttk.Label(
            main_frame, text="选择标签模板", font=("Arial", 14, "bold")
        )
        title_label.pack(pady=(0, 20))

        # 模板选择变量
        self.template_choice = tk.StringVar(value="常规")

        # 模板选择框架
        template_frame = ttk.LabelFrame(main_frame, text="模板类型", padding="15")
        template_frame.pack(fill=tk.X, pady=(0, 20))

        # 三个模板选项
        templates = [
            ("常规", "适用于普通包装标签"),
            ("分盒", "适用于分盒包装标签"),
            ("套盒", "适用于套盒包装标签")
        ]

        for i, (template_name, description) in enumerate(templates):
            radio = ttk.Radiobutton(
                template_frame, 
                text=f"{template_name} - {description}",
                variable=self.template_choice,
                value=template_name
            )
            radio.pack(anchor=tk.W, pady=5)

        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))

        self.selected_template = None

        def confirm_template():
            self.selected_template = self.template_choice.get()
            dialog.destroy()

        def cancel_template():
            self.selected_template = None
            dialog.destroy()

        # 确认和取消按钮
        ttk.Button(button_frame, text="确认", command=confirm_template).pack(side=tk.RIGHT, padx=(10, 0))
        ttk.Button(button_frame, text="取消", command=cancel_template).pack(side=tk.RIGHT)

        # 等待对话框关闭
        dialog.wait_window()
        return self.selected_template

    def start_generation_workflow(self):
        """开始生成工作流：先选择模板，再设置参数"""
        if not self.current_data:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return

        # 步骤1: 选择模板
        template_choice = self.show_template_selection_dialog()
        if not template_choice:
            return  # 用户取消选择

        # 保存模板选择
        self.selected_main_template = template_choice

        # 步骤2: 设置参数（根据模板调整参数界面）
        self.show_parameters_dialog_for_template(template_choice)

    def show_parameters_dialog_for_template(self, template_type):
        """根据模板类型显示对应的参数设置对话框"""
        if template_type == "常规":
            self.show_parameters_dialog()
        elif template_type == "分盒":
            self.show_fenhe_parameters_dialog()  # 分盒模板专用对话框
        elif template_type == "套盒":
            self.show_taohebox_parameters_dialog()  # 套盒模板专用对话框

    def generate_multi_level_pdfs(self):
        """生成多级标签PDF"""
        if not self.current_data or not self.packaging_params:
            messagebox.showwarning("警告", "缺少必要数据或参数")
            return

        # 使用已选择的模板
        template_choice = getattr(self, 'selected_main_template', '常规')

        try:
            self.status_var.set(f"🔄 正在生成{template_choice}模板PDF...")
            self.info_text.insert(tk.END, f"\n开始生成{template_choice}模板PDF...\n")
            self.root.update()

            # 选择输出目录
            output_dir = filedialog.askdirectory(
                title="选择输出目录", initialdir=os.path.expanduser("~/Desktop")
            )

            if output_dir:
                # 创建PDF生成器
                generator = PDFGenerator()
                
                # 根据模板选择调用不同的生成方法
                if template_choice == "常规":
                    generated_files = generator.create_multi_level_pdfs(
                        self.current_data, self.packaging_params, output_dir, self.current_file
                    )
                elif template_choice == "分盒":
                    generated_files = generator.create_fenhe_multi_level_pdfs(
                        self.current_data, self.packaging_params, output_dir, self.current_file
                    )
                elif template_choice == "套盒":
                    generated_files = generator.create_taohebox_multi_level_pdfs(
                        self.current_data, self.packaging_params, output_dir, self.current_file
                    )

                self.status_var.set(f"✅ {template_choice}模板PDF生成成功!")

                # 显示生成结果
                result_text = "\n✅ 生成完成! 文件列表:\n"
                for label_type, file_path in generated_files.items():
                    result_text += f"  - {label_type}: {Path(file_path).name}\n"

                folder_name = (
                    f"{self.current_data['客户编码']}+{self.current_data['主题']}+标签"
                )
                result_text += (
                    f"\n📁 保存目录: {os.path.join(output_dir, folder_name)}\n"
                )

                self.info_text.insert(tk.END, result_text)

                # 询问是否打开文件夹
                if messagebox.askyesno(
                    "生成成功",
                    f"多级标签PDF已生成!\n\n保存目录: {folder_name}\n\n是否打开文件夹？",
                ):
                    import subprocess
                    import platform

                    folder_path = os.path.join(output_dir, folder_name)
                    if platform.system() == "Darwin":  # macOS
                        subprocess.run(["open", folder_path])
                    elif platform.system() == "Windows":  # Windows
                        os.startfile(folder_path)

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
        if file_path.lower().endswith((".xlsx", ".xls")):
            # 延迟处理文件，等GUI完全加载
            root.after(500, lambda: app.process_file(file_path))

    root.mainloop()


if __name__ == "__main__":
    main()
