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

        # 是否有小箱选择框架
        small_box_frame = ttk.LabelFrame(main_frame, text="包装类型选择", padding="15")
        small_box_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.main_app.has_small_box_var = tk.StringVar(value="有小箱")
        
        # 居中布局的框架
        small_box_container = ttk.Frame(small_box_frame)
        small_box_container.pack(expand=True)
        
        has_small_box_radio = ttk.Radiobutton(
            small_box_container,
            text="有小箱（三级包装）",
            variable=self.main_app.has_small_box_var,
            value="有小箱",
            command=self.on_small_box_choice_changed
        )
        has_small_box_radio.grid(row=0, column=0, sticky=tk.W, pady=5)

        no_small_box_radio = ttk.Radiobutton(
            small_box_container,
            text="无小箱（二级包装）",
            variable=self.main_app.has_small_box_var,
            value="无小箱",
            command=self.on_small_box_choice_changed
        )
        no_small_box_radio.grid(row=0, column=1, sticky=tk.W, padx=(20, 0), pady=5)
        
        # 盒标类型选择框架
        box_label_frame = ttk.LabelFrame(main_frame, text="盒标类型选择", padding="15")
        box_label_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.main_app.has_box_label_var = tk.StringVar(value="无盒标")
        
        # 居中布局的框架
        box_label_container = ttk.Frame(box_label_frame)
        box_label_container.pack(expand=True)
        
        no_box_label_radio = ttk.Radiobutton(
            box_label_container,
            text="无盒标",
            variable=self.main_app.has_box_label_var,
            value="无盒标"
        )
        no_box_label_radio.grid(row=0, column=0, sticky=tk.W, pady=5)

        has_box_label_radio = ttk.Radiobutton(
            box_label_container,
            text="有盒标",
            variable=self.main_app.has_box_label_var,
            value="有盒标"
        )
        has_box_label_radio.grid(row=0, column=1, sticky=tk.W, padx=(20, 0), pady=5)
        
        # 参数输入框架
        params_frame = ttk.LabelFrame(main_frame, text="包装参数", padding="15")
        params_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 保存params_frame引用以便后续使用
        self.params_frame = params_frame

        # 张/盒输入框（预填充Excel提取的值）
        ttk.Label(params_frame, text="张/盒:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        self.main_app.pieces_per_box_var = tk.StringVar()
        pieces_per_box_entry = ttk.Entry(
            params_frame, textvariable=self.main_app.pieces_per_box_var, width=15
        )
        pieces_per_box_entry.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        # 预填充从Excel提取的值
        pieces_per_box_value = self.main_app.current_data.get('张/盒', '')
        if pieces_per_box_value and pieces_per_box_value != 'N/A':
            self.main_app.pieces_per_box_var.set(str(pieces_per_box_value))

        # 第二个参数（根据选择动态调整）
        self.second_param_label = ttk.Label(params_frame, text="盒/小箱:")
        self.second_param_label.grid(row=1, column=0, sticky=tk.W, pady=5)
        self.main_app.boxes_per_small_box_var = tk.StringVar()
        self.second_param_entry = ttk.Entry(
            params_frame, textvariable=self.main_app.boxes_per_small_box_var, width=15
        )
        self.second_param_entry.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=5)

        # 第三个参数（只在有小箱时显示）
        self.third_param_label = ttk.Label(params_frame, text="小箱/大箱:")
        self.third_param_label.grid(row=2, column=0, sticky=tk.W, pady=5)
        self.main_app.small_boxes_per_large_box_var = tk.StringVar()
        self.third_param_entry = ttk.Entry(
            params_frame, textvariable=self.main_app.small_boxes_per_large_box_var, width=15
        )
        self.third_param_entry.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=5)

        # 中文名称输入
        ttk.Label(params_frame, text="中文名称:").grid(
            row=3, column=0, sticky=tk.W, pady=5
        )
        self.main_app.chinese_name_var = tk.StringVar()
        chinese_name_entry = ttk.Entry(
            params_frame, textvariable=self.main_app.chinese_name_var, width=15
        )
        chinese_name_entry.grid(row=3, column=1, sticky=tk.W, padx=(10, 0), pady=5)

        # 序列号字体大小输入
        ttk.Label(params_frame, text="序列号字体大小:").grid(
            row=4, column=0, sticky=tk.W, pady=5
        )
        self.main_app.serial_font_size_var = tk.StringVar(value="10")
        serial_font_size_entry = ttk.Entry(
            params_frame, textvariable=self.main_app.serial_font_size_var, width=15
        )
        serial_font_size_entry.grid(row=4, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        # 添加说明标签
        serial_font_help = ttk.Label(
            params_frame, 
            text="(建议范围: 6-14, 序列号较长时可调小)",
            font=('Arial', 8),
            foreground='gray'
        )
        serial_font_help.grid(row=4, column=2, sticky=tk.W, padx=(5, 0), pady=5)
        
        # 初始化时设置显示状态
        self.on_small_box_choice_changed()

        # 标签模版选择框架
        template_frame = ttk.LabelFrame(main_frame, text="标签模版选择", padding="15")
        template_frame.pack(fill=tk.X, pady=(0, 20))

        self.main_app.template_var = tk.StringVar(value="无纸卡备注")
        
        # 创建单选按钮
        template_radio1 = ttk.Radiobutton(
            template_frame,
            text="无纸卡备注",
            variable=self.main_app.template_var,
            value="无纸卡备注"
        )
        template_radio1.grid(row=0, column=0, sticky=tk.W, pady=5)

        template_radio2 = ttk.Radiobutton(
            template_frame,
            text="有纸卡备注",
            variable=self.main_app.template_var,
            value="有纸卡备注"
        )
        template_radio2.grid(row=0, column=1, sticky=tk.W, padx=(20, 0), pady=5)

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
        # 获取小箱选择
        has_small_box = self.main_app.has_small_box_var.get() == "有小箱"
        
        # 获取参数值
        boxes_per_small_box_str = self.main_app.boxes_per_small_box_var.get().strip()
        small_boxes_per_large_box_str = self.main_app.small_boxes_per_large_box_var.get().strip()
        
        # 获取张/盒参数值
        pieces_per_box_str = self.main_app.pieces_per_box_var.get().strip()
        
        # 检查张/盒空值
        if not pieces_per_box_str:
            messagebox.showerror("参数错误", "请输入'张/盒'参数")
            return
            
        try:
            pieces_per_box = int(pieces_per_box_str)
        except ValueError:
            messagebox.showerror("参数错误", "'张/盒'必须为有效整数\n\n正确格式示例：300")
            return
        
        # 检查张/盒是否有效
        if pieces_per_box <= 0:
            messagebox.showerror("参数错误", "'张/盒'必须为正整数\n\n当前值：{}".format(pieces_per_box))
            return
        
        # 检查张/盒不能超过总张数
        total_pieces = int(self.main_app.current_data.get('总张数', 0))
        if pieces_per_box > total_pieces:
            messagebox.showerror("参数错误", "'张/盒'不能超过总张数\n\n当前设置：{} 张/盒\n总张数：{} 张\n\n请输入不超过{}的值".format(pieces_per_box, total_pieces, total_pieces))
            return
        
        # 检查空值
        if not boxes_per_small_box_str:
            if has_small_box:
                messagebox.showerror("参数错误", "请输入'盒/小箱'参数")
            else:
                messagebox.showerror("参数错误", "请输入'盒/箱'参数")
            return
        if has_small_box and not small_boxes_per_large_box_str:
            messagebox.showerror("参数错误", "请输入'小箱/大箱'参数")
            return
        
        try:
            # 尝试转换为数字
            boxes_per_small_box = int(boxes_per_small_box_str)
            if has_small_box:
                small_boxes_per_large_box = int(small_boxes_per_large_box_str)
            else:
                # 无小箱时，设置为1（表示跳过小箱层级）
                small_boxes_per_large_box = 1
        except ValueError:
            if has_small_box:
                messagebox.showerror("参数错误", "请输入有效的整数\n\n正确格式示例：3、2")
            else:
                messagebox.showerror("参数错误", "请输入有效的整数\n\n正确格式示例：6")
            return
        
        # 检查负数和0
        if boxes_per_small_box <= 0:
            if has_small_box:
                messagebox.showerror("参数错误", "'盒/小箱'必须为正整数\n\n当前值：{}".format(boxes_per_small_box))
            else:
                messagebox.showerror("参数错误", "'盒/箱'必须为正整数\n\n当前值：{}".format(boxes_per_small_box))
            return
        if has_small_box and small_boxes_per_large_box <= 0:
            messagebox.showerror("参数错误", "'小箱/大箱'必须为正整数\n\n当前值：{}".format(small_boxes_per_large_box))
            return
        
            
        # 获取中文名称
        chinese_name = self.main_app.chinese_name_var.get().strip()
        
        # 检查中文名称是否为空
        if not chinese_name:
            messagebox.showerror("参数错误", "请输入'中文名称'")
            return
        
        # 获取序列号字体大小
        serial_font_size_str = self.main_app.serial_font_size_var.get().strip()
        if not serial_font_size_str:
            messagebox.showerror("参数错误", "请输入'序列号字体大小'")
            return
        
        try:
            serial_font_size = int(serial_font_size_str)
            if serial_font_size < 6 or serial_font_size > 14:
                messagebox.showerror("参数错误", "序列号字体大小必须在6-14之间\n\n当前值：{}".format(serial_font_size))
                return
        except ValueError:
            messagebox.showerror("参数错误", "请输入有效的序列号字体大小\n\n正确格式示例：10")
            return
        
        # 获取盒标选择
        has_box_label = self.main_app.has_box_label_var.get() == "有盒标"
        
        # 参数验证通过，设置参数
        self.main_app.packaging_params = {
            "张/盒": pieces_per_box,
            "盒/小箱": boxes_per_small_box,
            "小箱/大箱": small_boxes_per_large_box,
            "选择外观": "外观一",  # 分盒模板固定使用外观一
            "标签模版": self.main_app.template_var.get(),
            "中文名称": self.main_app.chinese_name_var.get(),
            "是否有小箱": has_small_box,
            "序列号字体大小": serial_font_size,
            "是否有盒标": has_box_label,
        }

        dialog.destroy()
        self.main_app.generate_multi_level_pdfs()
    
    def on_small_box_choice_changed(self):
        """处理小箱选择变化"""
        has_small_box = self.main_app.has_small_box_var.get() == "有小箱"
        
        if has_small_box:
            # 有小箱：三级包装
            self.second_param_label.config(text="盒/小箱:")
            self.third_param_label.grid()
            self.third_param_entry.grid()
        else:
            # 无小箱：二级包装
            self.second_param_label.config(text="盒/箱:")
            self.third_param_label.grid_remove()
            self.third_param_entry.grid_remove()


# 创建全局实例供split_box模板使用
split_box_ui_dialog = None

def get_split_box_ui_dialog(main_app):
    """获取分盒模板UI对话框实例（单例模式）"""
    global split_box_ui_dialog
    if split_box_ui_dialog is None:
        split_box_ui_dialog = SplitBoxUIDialog(main_app)
    return split_box_ui_dialog