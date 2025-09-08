"""
套盒模板UI对话框
专门处理套盒模板的参数设置界面
"""

import tkinter as tk
from tkinter import ttk, messagebox


class NestedBoxUIDialog:
    """套盒模板UI对话框处理类"""
    
    def __init__(self, main_app):
        """
        初始化套盒模板UI对话框
        
        Args:
            main_app: 主GUI应用程序实例
        """
        self.main_app = main_app
    
    def show_parameters_dialog(self):
        """显示套盒模板的参数设置对话框（无外观选择）"""
        if not self.main_app.current_data:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return

        # 创建对话框
        dialog = tk.Toplevel(self.main_app.root)
        dialog.title("套盒模板 - 包装参数设置")
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
        self.main_app.pieces_per_box_var = tk.StringVar()
        pieces_per_box_entry = ttk.Entry(
            params_frame, textvariable=self.main_app.pieces_per_box_var, width=15
        )
        pieces_per_box_entry.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=5)

        # 盒/小箱输入
        ttk.Label(params_frame, text="盒/小箱:").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        self.main_app.boxes_per_small_box_var = tk.StringVar()
        boxes_per_small_box_entry = ttk.Entry(
            params_frame, textvariable=self.main_app.boxes_per_small_box_var, width=15
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

        data_text = f"客户编码: {self.main_app.current_data['客户编码']}\n"
        data_text += f"主题: {self.main_app.current_data['主题']}\n"
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
        self._auto_resize_and_center_dialog(dialog, scrollable_frame)

    def confirm_parameters(self, dialog):
        """确认套盒模板参数并生成PDF"""
        # 获取参数值
        pieces_per_box_str = self.main_app.pieces_per_box_var.get().strip()
        boxes_per_small_box_str = self.main_app.boxes_per_small_box_var.get().strip()
        small_boxes_per_large_box_str = self.main_app.small_boxes_per_large_box_var.get().strip()
        
        # 检查空值
        if not pieces_per_box_str:
            messagebox.showerror("参数错误", "请输入“张/盒”参数")
            return
        if not boxes_per_small_box_str:
            messagebox.showerror("参数错误", "请输入“盒/小箱”参数")
            return
        if not small_boxes_per_large_box_str:
            messagebox.showerror("参数错误", "请输入“小箱/大箱”参数")
            return
        
        try:
            # 尝试转换为数字
            pieces_per_box = int(pieces_per_box_str)
            boxes_per_small_box = int(boxes_per_small_box_str)
            small_boxes_per_large_box = int(small_boxes_per_large_box_str)
        except ValueError:
            messagebox.showerror("参数错误", "请输入有效的整数\n\n正确格式示例：300、6、2")
            return
        
        # 检查负数和0
        if pieces_per_box <= 0:
            messagebox.showerror("参数错误", "“张/盒”必须为正整数\n\n当前值：{}".format(pieces_per_box))
            return
        if boxes_per_small_box <= 0:
            messagebox.showerror("参数错误", "“盒/小箱”必须为正整数\n\n当前值：{}".format(boxes_per_small_box))
            return
        if small_boxes_per_large_box <= 0:
            messagebox.showerror("参数错误", "“小箱/大箱”必须为正整数\n\n当前值：{}".format(small_boxes_per_large_box))
            return
        
        # 参数验证通过，设置参数
        self.main_app.packaging_params = {
            "张/盒": pieces_per_box,
            "盒/小箱": boxes_per_small_box,
            "小箱/大箱": small_boxes_per_large_box,
            "选择外观": "外观一",  # 套盒模板固定使用外观一，但实际不使用
        }

        dialog.destroy()
        self.main_app.generate_multi_level_pdfs()

    def _auto_resize_and_center_dialog(self, dialog, content_frame):
        """自动调整对话框大小并居中显示，完全基于内容自适应"""
        try:
            # 多次更新确保所有组件都已完全渲染
            for _ in range(3):
                dialog.update_idletasks()
                content_frame.update_idletasks()
            
            # 获取内容的实际所需尺寸
            content_width = content_frame.winfo_reqwidth()
            content_height = content_frame.winfo_reqheight()
            
            # 添加必要的边距：滚动条、对话框边框、标题栏等
            padding_width = 60   # 减少左右边距
            padding_height = 80   # 减少上下边距，让对话框更紧凑
            
            # 计算对话框所需的实际尺寸
            required_width = content_width + padding_width
            required_height = content_height + padding_height
            
            # 获取屏幕尺寸，确保不会超出屏幕
            screen_width = dialog.winfo_screenwidth()
            screen_height = dialog.winfo_screenheight()
            
            # 最终尺寸：完全基于内容，但不超过屏幕90%
            final_width = min(required_width, int(screen_width * 0.9))
            final_height = min(required_height, int(screen_height * 0.9))
            
            # 计算居中位置
            x = (screen_width - final_width) // 2
            y = (screen_height - final_height) // 2
            
            # 设置对话框几何形状
            dialog.geometry(f"{final_width}x{final_height}+{x}+{y}")
            
            print(f"✅ 完全自适应调整: {final_width}x{final_height}")
            print(f"   内容尺寸: {content_width}x{content_height}")
            print(f"   边距: {padding_width}x{padding_height}")
            
        except Exception as e:
            print(f"⚠️ 自适应调整失败: {e}")
            # 备用方案：让系统自动计算
            dialog.update_idletasks()
            dialog.geometry("")  # 清空几何设置，让Tkinter自动调整
            dialog.update_idletasks()
            
            # 获取自动调整后的尺寸并居中
            width = dialog.winfo_width()
            height = dialog.winfo_height()
            x = (dialog.winfo_screenwidth() - width) // 2
            y = (dialog.winfo_screenheight() - height) // 2
            dialog.geometry(f"{width}x{height}+{x}+{y}")


# 全局变量用于单例模式
nested_box_ui_dialog = None


def get_nested_box_ui_dialog(main_app):
    """获取套盒模板UI对话框实例（单例模式）"""
    global nested_box_ui_dialog
    if nested_box_ui_dialog is None:
        nested_box_ui_dialog = NestedBoxUIDialog(main_app)
    return nested_box_ui_dialog