"""
常规模板UI对话框
专门处理常规模板的参数设置界面
"""

import tkinter as tk
from tkinter import ttk, messagebox


class RegularUIDialog:
    """常规模板UI对话框处理类"""
    
    def __init__(self, main_app):
        """
        初始化常规模板UI对话框
        
        Args:
            main_app: 主GUI应用程序实例
        """
        self.main_app = main_app
    
    def show_parameters_dialog(self):
        """显示常规模板参数设置对话框"""
        if not self.main_app.current_data:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return

        # 创建对话框
        dialog = tk.Toplevel(self.main_app.root)
        dialog.title("常规模板 - 包装参数设置")
        dialog.transient(self.main_app.root)
        dialog.grab_set()
        
        # 设置最小尺寸，确保能显示全部内容
        dialog.minsize(550, 450)
        dialog.resizable(True, True)
        
        # 先设置一个初始尺寸
        dialog.geometry("580x480")

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

        # 创建一个居中的内容容器
        center_container = ttk.Frame(main_frame)
        center_container.pack(expand=True, fill=tk.BOTH)
        
        # 标题
        title_label = ttk.Label(
            center_container, text="常规模板参数设置", font=("Arial", 14, "bold")
        )
        title_label.pack(pady=(0, 15))

        # 创建自适应布局的主容器
        main_content_frame = ttk.Frame(center_container)
        main_content_frame.pack(pady=(0, 10), padx=10, fill=tk.BOTH, expand=True)
        
        # 使用grid布局实现响应式设计
        main_content_frame.grid_columnconfigure(0, weight=1)
        main_content_frame.grid_columnconfigure(1, weight=1)
        main_content_frame.grid_rowconfigure(0, weight=0)
        main_content_frame.grid_rowconfigure(1, weight=1)
        
        # 左侧列 - 参数输入和外观选择
        left_column = ttk.Frame(main_content_frame)
        left_column.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, 5))
        
        # 右侧列 - 当前数据
        right_column = ttk.Frame(main_content_frame)
        right_column.grid(row=0, column=1, sticky=tk.NSEW, padx=(5, 0))

        # 左侧：参数输入框架
        params_frame = ttk.LabelFrame(left_column, text="包装参数", padding="12")
        params_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 配置网格权重实现居中布局
        params_frame.grid_columnconfigure(0, weight=1)
        params_frame.grid_columnconfigure(2, weight=1)

        # 张/盒输入
        ttk.Label(params_frame, text="张/盒:", font=("Arial", 11)).grid(
            row=0, column=0, sticky=tk.E, pady=10, padx=(0, 15)
        )
        self.main_app.pieces_per_box_var = tk.StringVar()
        pieces_per_box_entry = ttk.Entry(
            params_frame, textvariable=self.main_app.pieces_per_box_var, width=20, font=("Arial", 11)
        )
        pieces_per_box_entry.grid(row=0, column=1, sticky=tk.W+tk.E, pady=10)

        # 盒/小箱输入
        ttk.Label(params_frame, text="盒/小箱:", font=("Arial", 11)).grid(
            row=1, column=0, sticky=tk.E, pady=10, padx=(0, 15)
        )
        self.main_app.boxes_per_small_box_var = tk.StringVar()
        boxes_per_small_box_entry = ttk.Entry(
            params_frame, textvariable=self.main_app.boxes_per_small_box_var, width=20, font=("Arial", 11)
        )
        boxes_per_small_box_entry.grid(row=1, column=1, sticky=tk.W+tk.E, pady=10)

        # 小箱/大箱输入
        ttk.Label(params_frame, text="小箱/大箱:", font=("Arial", 11)).grid(
            row=2, column=0, sticky=tk.E, pady=10, padx=(0, 15)
        )
        self.main_app.small_boxes_per_large_box_var = tk.StringVar()
        small_boxes_entry = ttk.Entry(
            params_frame, textvariable=self.main_app.small_boxes_per_large_box_var, width=20, font=("Arial", 11)
        )
        small_boxes_entry.grid(row=2, column=1, sticky=tk.W+tk.E, pady=10)

        # 左侧：外观选择框架
        appearance_frame = ttk.LabelFrame(left_column, text="盒标外观选择", padding="12")
        appearance_frame.pack(fill=tk.X)

        self.main_app.appearance_var = tk.StringVar(value="外观一")
        
        # 居中布局的框架
        radio_container = ttk.Frame(appearance_frame)
        radio_container.pack(expand=True)
        
        appearance_radio1 = ttk.Radiobutton(
            radio_container,
            text="外观一",
            variable=self.main_app.appearance_var,
            value="外观一"
        )
        appearance_radio1.pack(side=tk.LEFT, padx=(0, 30))

        appearance_radio2 = ttk.Radiobutton(
            radio_container,
            text="外观二",
            variable=self.main_app.appearance_var,
            value="外观二"
        )
        appearance_radio2.pack(side=tk.LEFT)

        # 右侧：当前数据显示
        info_frame = ttk.LabelFrame(right_column, text="当前数据", padding="12")
        info_frame.pack(fill=tk.BOTH, expand=True)
        
        # 使用网格布局实现更好的对齐
        info_frame.grid_columnconfigure(1, weight=1)

        # 客户编码
        ttk.Label(info_frame, text="客户名称编码:", font=("Arial", 11)).grid(
            row=0, column=0, sticky=tk.E, pady=12, padx=(0, 15)
        )
        ttk.Label(info_frame, text=self.main_app.current_data['客户名称编码'], font=("Arial", 11, "bold")).grid(
            row=0, column=1, sticky=tk.W, pady=12
        )
        
        # 主题
        ttk.Label(info_frame, text="标签名称:", font=("Arial", 11)).grid(
            row=1, column=0, sticky=tk.NE, pady=12, padx=(0, 15)  # 改为NE，顶部对齐
        )
        theme_label = ttk.Label(info_frame, text=self.main_app.current_data['标签名称'], font=("Arial", 11, "bold"), wraplength=200, justify=tk.LEFT)
        theme_label.grid(row=1, column=1, sticky=tk.NW, pady=12)  # 改为NW，顶部左对齐
        
        # 总张数
        ttk.Label(info_frame, text="总张数:", font=("Arial", 11)).grid(
            row=2, column=0, sticky=tk.E, pady=12, padx=(0, 15)
        )
        ttk.Label(info_frame, text=self.main_app.current_data['总张数'], font=("Arial", 11, "bold")).grid(
            row=2, column=1, sticky=tk.W, pady=12
        )

        # 按钮框架 - 居中布局
        button_frame = ttk.Frame(center_container)
        button_frame.pack(pady=(15, 0))
        
        # 创建居中容器
        button_container = ttk.Frame(button_frame)
        button_container.pack(expand=True)

        # 确认按钮
        confirm_btn = ttk.Button(
            button_container,
            text="确认生成",
            command=lambda: self.confirm_parameters(dialog),
        )
        confirm_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 取消按钮
        cancel_btn = ttk.Button(
            button_container, text="取消", command=dialog.destroy
        )
        cancel_btn.pack(side=tk.LEFT)

        # 自动调整对话框大小
        self.main_app._auto_resize_and_center_dialog(dialog, scrollable_frame)

    def confirm_parameters(self, dialog):
        """确认常规模板参数并生成PDF"""
        # 获取参数值
        pieces_per_box_str = self.main_app.pieces_per_box_var.get().strip()
        boxes_per_small_box_str = self.main_app.boxes_per_small_box_var.get().strip()
        small_boxes_per_large_box_str = self.main_app.small_boxes_per_large_box_var.get().strip()
        
        # 检查空值
        if not pieces_per_box_str:
            messagebox.showerror("参数错误", "请输入'张/盒'参数")
            return
        if not boxes_per_small_box_str:
            messagebox.showerror("参数错误", "请输入'盒/小箱'参数")
            return
        if not small_boxes_per_large_box_str:
            messagebox.showerror("参数错误", "请输入'小箱/大箱'参数")
            return
        
        try:
            # 尝试转换为数字
            pieces_per_box = int(pieces_per_box_str)
            boxes_per_small_box = int(boxes_per_small_box_str)
            small_boxes_per_large_box = int(small_boxes_per_large_box_str)
        except ValueError:
            messagebox.showerror("参数错误", "请输入有效的整数\n\n正确格式示例：300、2、5")
            return
        
        # 检查负数和0
        if pieces_per_box <= 0:
            messagebox.showerror("参数错误", "'张/盒'必须为正整数\n\n当前值：{}".format(pieces_per_box))
            return
        if boxes_per_small_box <= 0:
            messagebox.showerror("参数错误", "'盒/小箱'必须为正整数\n\n当前值：{}".format(boxes_per_small_box))
            return
        if small_boxes_per_large_box <= 0:
            messagebox.showerror("参数错误", "'小箱/大箱'必须为正整数\n\n当前值：{}".format(small_boxes_per_large_box))
            return
        
        # 检查张/盒不能超过总张数
        total_pieces = int(self.main_app.current_data.get('总张数', 0))
        if pieces_per_box > total_pieces:
            messagebox.showerror("参数错误", "'张/盒'不能超过总张数\n\n当前设置：{} 张/盒\n总张数：{} 张\n\n请输入不超过 {} 的值".format(pieces_per_box, total_pieces, total_pieces))
            return
        
        # 参数验证通过，设置参数
        self.main_app.packaging_params = {
            "张/盒": pieces_per_box,
            "盒/小箱": boxes_per_small_box,
            "小箱/大箱": small_boxes_per_large_box,
            "选择外观": self.main_app.appearance_var.get(),
        }

        dialog.destroy()
        self.main_app.generate_multi_level_pdfs()


# 创建全局实例供regular模板使用
regular_ui_dialog = None

def get_regular_ui_dialog(main_app):
    """获取常规模板UI对话框实例（单例模式）"""
    global regular_ui_dialog
    if regular_ui_dialog is None:
        regular_ui_dialog = RegularUIDialog(main_app)
    return regular_ui_dialog