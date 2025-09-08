"""
分盒模板UI对话框
专门处理分盒模板的参数设置界面
"""

import tkinter as tk
from tkinter import ttk, messagebox


class SplitBoxUIDialog:
    """分盒模板UI对话框处理类"""
    
    def __init__(self, main_app):
        """
        初始化分盒模板UI对话框
        
        Args:
            main_app: 主GUI应用程序实例
        """
        self.main_app = main_app
    
    def show_parameters_dialog(self):
        """显示分盒模板的参数设置对话框（无外观选择）"""
        if not self.main_app.current_data:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return

        # 创建对话框
        dialog = tk.Toplevel(self.main_app.root)
        dialog.title("分盒模板 - 包装参数设置")
        dialog.transient(self.main_app.root)
        dialog.grab_set()
        
        # 设置最小尺寸
        dialog.minsize(480, 380)
        
        # 先设置一个初始尺寸
        dialog.geometry("500x420")

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
        main_frame = ttk.Frame(scrollable_frame, padding="30")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 配置主框架的网格权重，让内容能更好地填充空间
        main_frame.grid_columnconfigure(0, weight=1)

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
        self.main_app.pieces_per_box_var = tk.StringVar()
        pieces_per_box_entry = ttk.Entry(
            params_frame, textvariable=self.main_app.pieces_per_box_var, width=15
        )
        pieces_per_box_entry.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=5)

        # 盒/小箱输入 - 分盒模板固定为1，不可修改
        ttk.Label(params_frame, text="盒/小箱:").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        self.main_app.boxes_per_small_box_var = tk.StringVar(value="1")
        boxes_per_small_box_entry = ttk.Entry(
            params_frame, textvariable=self.main_app.boxes_per_small_box_var, width=15, state="disabled"
        )
        boxes_per_small_box_entry.grid(
            row=1, column=1, sticky=tk.W, padx=(10, 0), pady=5
        )

        # 小箱/大箱输入
        ttk.Label(params_frame, text="小箱/大箱:").grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        self.main_app.small_boxes_per_large_box_var = tk.StringVar()
        small_boxes_entry = ttk.Entry(
            params_frame, textvariable=self.main_app.small_boxes_per_large_box_var, width=15
        )
        small_boxes_entry.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=5)


        # 当前数据显示
        data_frame = ttk.LabelFrame(main_frame, text="当前数据", padding="15")
        data_frame.pack(fill=tk.X, pady=(0, 20))

        data_text = f"客户名称编码: {self.main_app.current_data['客户名称编码']}\n"
        data_text += f"标签名称: {self.main_app.current_data['标签名称']}\n"
        data_text += f"总张数: {self.main_app.current_data['总张数']}"

        data_label = ttk.Label(data_frame, text=data_text, font=("Consolas", 10))
        data_label.pack(anchor=tk.W)

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
        
        # 自适应大小和居中显示
        self.main_app._auto_resize_and_center_dialog(dialog, scrollable_frame)

    def confirm_parameters(self, dialog):
        """确认分盒模板参数并生成PDF"""
        # 获取参数值（第二个参数保持固定为1）
        pieces_per_box_str = self.main_app.pieces_per_box_var.get().strip()
        small_boxes_per_large_box_str = self.main_app.small_boxes_per_large_box_var.get().strip()
        
        # 检查空值
        if not pieces_per_box_str:
            messagebox.showerror("参数错误", "请输入“张/盒”参数")
            return
        if not small_boxes_per_large_box_str:
            messagebox.showerror("参数错误", "请输入“小箱/大箱”参数")
            return
        
        try:
            # 尝试转换为数字（第二个参数固定为1）
            pieces_per_box = int(pieces_per_box_str)
            boxes_per_small_box = 1  # 分盒模板固定为1
            small_boxes_per_large_box = int(small_boxes_per_large_box_str)
        except ValueError:
            messagebox.showerror("参数错误", "请输入有效的整数\n\n正确格式示例：300、2")
            return
        
        # 检查负数和0
        if pieces_per_box <= 0:
            messagebox.showerror("参数错误", "“张/盒”必须为正整数\n\n当前值：{}".format(pieces_per_box))
            return
        if small_boxes_per_large_box <= 0:
            messagebox.showerror("参数错误", ""小箱/大箱"必须为正整数\n\n当前值：{}".format(small_boxes_per_large_box))
            return
        
        # 检查张/盒不能超过总张数
        total_pieces = int(self.main_app.current_data.get('总张数', 0))
        if pieces_per_box > total_pieces:
            messagebox.showerror("参数错误", f""张/盒"不能超过总张数\n\n当前设置：{pieces_per_box} 张/盒\n总张数：{total_pieces} 张\n\n请输入不超过 {total_pieces} 的值")
            return
        
        # 参数验证通过，设置参数
        self.main_app.packaging_params = {
            "张/盒": pieces_per_box,
            "盒/小箱": boxes_per_small_box,
            "小箱/大箱": small_boxes_per_large_box,
            "选择外观": "外观一",  # 分盒模板固定使用外观一
        }

        dialog.destroy()
        self.main_app.generate_multi_level_pdfs()


# 创建全局实例供split_box模板使用
split_box_ui_dialog = None

def get_split_box_ui_dialog(main_app):
    """获取分盒模板UI对话框实例（单例模式）"""
    global split_box_ui_dialog
    if split_box_ui_dialog is None:
        split_box_ui_dialog = SplitBoxUIDialog(main_app)
    return split_box_ui_dialog