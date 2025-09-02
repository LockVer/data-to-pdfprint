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
            text="🔄 生成多级标签PDF",
            command=self.show_parameters_dialog,
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

            # 提取数据
            total_count = df.iloc[3, 5]  # 总张数

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

    def generate_multi_level_pdfs(self):
        """生成多级标签PDF"""
        if not self.current_data or not self.packaging_params:
            messagebox.showwarning("警告", "缺少必要数据或参数")
            return

        try:
            self.status_var.set("🔄 正在生成多级标签PDF...")
            self.info_text.insert(tk.END, "\n开始生成多级标签PDF...\n")
            self.root.update()

            # 选择输出目录
            output_dir = filedialog.askdirectory(
                title="选择输出目录", initialdir=os.path.expanduser("~/Desktop")
            )

            if output_dir:
                # 生成多级PDF
                generator = PDFGenerator()
                generated_files = generator.create_multi_level_pdfs(
                    self.current_data, self.packaging_params, output_dir
                )

                self.status_var.set("✅ 多级标签PDF生成成功!")

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
