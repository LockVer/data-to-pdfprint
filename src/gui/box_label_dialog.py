"""
盒标参数设置对话框

用于设置最小分盒张数、包装比例等参数
"""

import tkinter as tk
from tkinter import ttk, messagebox

# 避免循环导入，直接定义需要的样式类
class ModernColors:
    """现代化配色方案"""
    PRIMARY = '#2E7D8A'
    PRIMARY_LIGHT = '#4A9BAC'
    PRIMARY_DARK = '#1E5A65'
    SECONDARY = '#F8F9FA'
    ACCENT = '#FF6B6B'
    SUCCESS = '#51CF66'
    WARNING = '#FFD93D'
    ERROR = '#FF5252'
    WHITE = '#FFFFFF'
    LIGHT_GRAY = '#F1F3F4'
    GRAY = '#9AA0A6'
    DARK_GRAY = '#5F6368'
    BLACK = '#202124'
    CARD_BG = '#FFFFFF'
    SHADOW = '#E8EAED'
    BORDER = '#DADCE0'

class ModernFonts:
    """现代字体配置"""
    TITLE_LARGE = ('SF Pro Display', 28, 'bold')
    TITLE = ('SF Pro Display', 20, 'bold')
    SUBTITLE = ('SF Pro Display', 16, 'bold')
    BODY_LARGE = ('SF Pro Text', 14)
    BODY = ('SF Pro Text', 12)
    BODY_SMALL = ('SF Pro Text', 11)
    BUTTON = ('SF Pro Text', 12, 'bold')
    CODE = ('SF Mono', 11)

class ModernButton:
    """现代化按钮组件"""
    
    def __init__(self, parent, text, command=None, style='primary', width=200, height=45):
        self.frame = tk.Frame(parent, bg=parent['bg'] if hasattr(parent, '__getitem__') else ModernColors.SECONDARY)
        self.command = command
        self.style = style
        
        # 按钮样式配置
        styles = {
            'primary': {
                'bg': ModernColors.PRIMARY,
                'fg': ModernColors.WHITE,
                'hover_bg': ModernColors.PRIMARY_LIGHT,
                'active_bg': ModernColors.PRIMARY_DARK
            },
            'secondary': {
                'bg': ModernColors.LIGHT_GRAY,
                'fg': ModernColors.DARK_GRAY,
                'hover_bg': ModernColors.GRAY,
                'active_bg': ModernColors.BORDER
            },
            'success': {
                'bg': ModernColors.SUCCESS,
                'fg': ModernColors.WHITE,
                'hover_bg': '#45B85C',
                'active_bg': '#3DA85A'
            }
        }
        
        self.colors = styles.get(style, styles['primary'])
        
        # 创建按钮
        self.button = tk.Button(
            self.frame,
            text=text,
            command=self._on_click,
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            font=ModernFonts.BUTTON,
            relief='flat',
            bd=0,
            cursor='hand2',
            padx=20,
            pady=10,
            width=width//10,
            height=height//25
        )
        self.button.pack(fill='both', expand=True)
        
        # 绑定悬停效果
        self.button.bind('<Enter>', self._on_hover)
        self.button.bind('<Leave>', self._on_leave)
        self.button.bind('<Button-1>', self._on_press)
        self.button.bind('<ButtonRelease-1>', self._on_release)
    
    def _on_click(self):
        if self.command:
            self.command()
    
    def _on_hover(self, _event):
        self.button.config(bg=self.colors['hover_bg'])
    
    def _on_leave(self, _event):
        self.button.config(bg=self.colors['bg'])
    
    def _on_press(self, _event):
        self.button.config(bg=self.colors['active_bg'])
    
    def _on_release(self, _event):
        self.button.config(bg=self.colors['hover_bg'])
    
    def pack(self, **kwargs):
        self.frame.pack(**kwargs)
    
    def config(self, **kwargs):
        if 'state' in kwargs:
            self.button.config(state=kwargs['state'])
            if kwargs['state'] == 'disabled':
                self.button.config(bg=ModernColors.GRAY, cursor='arrow')
            else:
                self.button.config(bg=self.colors['bg'], cursor='hand2')

class ModernCard:
    """现代化卡片组件"""
    
    def __init__(self, parent, title=None, padding=20):
        self.frame = tk.Frame(
            parent,
            bg=ModernColors.CARD_BG,
            relief='flat',
            bd=1,
            highlightbackground=ModernColors.BORDER,
            highlightthickness=1
        )
        
        self.content_frame = tk.Frame(self.frame, bg=ModernColors.CARD_BG)
        self.content_frame.pack(fill='both', expand=True, padx=padding, pady=padding)
        
        if title:
            title_label = tk.Label(
                self.content_frame,
                text=title,
                font=ModernFonts.SUBTITLE,
                bg=ModernColors.CARD_BG,
                fg=ModernColors.BLACK,
                anchor='w'
            )
            title_label.pack(fill='x', pady=(0, 15))
    
    def pack(self, **kwargs):
        self.frame.pack(**kwargs)
    
    def get_content_frame(self):
        return self.content_frame

class BoxLabelConfigDialog:
    """盒标配置对话框"""
    
    def __init__(self, parent, initial_config=None, template_type="regular", is_division_box=False):
        """
        初始化对话框
        
        Args:
            parent: 父窗口
            initial_config: 初始配置字典
            template_type: 模版类型 ("regular" 或 "set_box")
        """
        self.parent = parent
        self.result = None
        self.template_type = template_type  # 保存模版类型
        self.is_division_box = is_division_box  # 是否为分盒模版
        self.dialog = tk.Toplevel(parent)
        self.setup_dialog()
        self.create_widgets()
        
        # 设置初始值
        if initial_config:
            self.set_config(initial_config)
    
    def get_template_config(self):
        """根据模版类型获取参数配置"""
        if self.template_type == "set_box":
            return {
                "title": "套盒标生成参数设置",
                "params": [
                    ("最小分盒套数:", "min_set_count", "每套包装的张数", "30"),
                    ("几盒为一套:", "boxes_per_set", "多少盒组成一套", "3"),
                    ("几盒入一小箱:", "boxes_per_inner_case", "小箱可以装多少盒", "6"),
                    ("几套入一大箱:", "sets_per_outer_case", "大箱可以装多少套", "2")
                ]
            }
        else:  # regular 常规模版
            return {
                "title": "盒标生成参数设置",
                "params": [
                    ("最小分盒张数:", "min_box_count", "每盒包装的最少张数", "10"),
                    ("盒/小箱比例:", "box_per_inner_case", "每个小箱包装的盒数", "5"),
                    ("小箱/大箱比例:", "inner_case_per_outer_case", "每个大箱包装的小箱数", "4")
                ]
            }
        
    def setup_dialog(self):
        """设置对话框属性"""
        template_config = self.get_template_config()
        self.dialog.title(template_config["title"])
        self.dialog.geometry("500x600")  # 进一步增加高度
        self.dialog.configure(bg=ModernColors.SECONDARY)
        self.dialog.resizable(True, True)  # 允许调整大小
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # 绑定关闭事件
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel)
        
        # 绑定键盘快捷键
        self.dialog.bind('<Return>', lambda e: self.ok())
        self.dialog.bind('<Escape>', lambda e: self.cancel())
        
        # 窗口居中
        self.center_dialog()
        
    def center_dialog(self):
        """对话框居中显示"""
        self.dialog.update_idletasks()
        width = 500
        height = 600  # 进一步增加高度
        x = self.parent.winfo_x() + (self.parent.winfo_width() // 2) - (width // 2)
        y = self.parent.winfo_y() + (self.parent.winfo_height() // 2) - (height // 2)
        self.dialog.geometry(f'{width}x{height}+{x}+{y}')
        
    def create_widgets(self):
        """创建对话框组件"""
        # 主容器
        main_frame = tk.Frame(self.dialog, bg=ModernColors.SECONDARY)
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 顶部按钮区域（临时解决方案）
        top_button_frame = tk.Frame(main_frame, bg=ModernColors.SECONDARY)
        top_button_frame.pack(fill='x', pady=(0, 10))
        
        # 取消按钮（顶部）
        tk.Button(
            top_button_frame,
            text="❌ 取消",
            command=self.cancel,
            font=ModernFonts.BUTTON,
            bg=ModernColors.LIGHT_GRAY,
            fg=ModernColors.DARK_GRAY,
            relief='flat',
            padx=15,
            pady=5
        ).pack(side='left')
        
        # 确定按钮（顶部）
        tk.Button(
            top_button_frame,
            text="✅ 确定",
            command=self.ok,
            font=ModernFonts.BUTTON,
            bg=ModernColors.SUCCESS,
            fg=ModernColors.WHITE,
            relief='flat',
            padx=15,
            pady=5
        ).pack(side='right')
        
        # 标题
        title_label = tk.Label(
            main_frame,
            text="📦 盒标生成参数设置",
            font=ModernFonts.TITLE,
            bg=ModernColors.SECONDARY,
            fg=ModernColors.BLACK
        )
        title_label.pack(pady=(10, 20))
        
        # 参数设置卡片
        config_card = ModernCard(main_frame, "参数配置")
        config_card.pack(fill='both', expand=True, pady=(0, 20))
        content = config_card.get_content_frame()
        
        # 动态创建参数输入框
        template_config = self.get_template_config()
        for label, field_name, description, default_value in template_config["params"]:
            self.create_input_row(content, label, field_name, description, default_value)
        
        # 预览信息区域 - 限制高度
        preview_card = ModernCard(main_frame, "生成预览")
        preview_card.pack(fill='x', pady=(0, 20))
        preview_content = preview_card.get_content_frame()
        
        self.preview_label = tk.Label(
            preview_content,
            text="请输入参数后查看生成预览",
            font=ModernFonts.BODY,
            bg=ModernColors.CARD_BG,
            fg=ModernColors.GRAY,
            justify='left',
            anchor='nw',
            wraplength=400,
            height=8  # 限制高度为8行
        )
        self.preview_label.pack(fill='x')
        
        # 按钮区域 - 使用简单的标准按钮
        button_frame = tk.Frame(main_frame, bg=ModernColors.SECONDARY)
        button_frame.pack(fill='x', pady=(30, 10))
        
        # 取消按钮 - 使用标准tkinter按钮确保显示
        self.cancel_btn = tk.Button(
            button_frame,
            text="❌ 取消",
            command=self.cancel,
            font=ModernFonts.BUTTON,
            bg=ModernColors.LIGHT_GRAY,
            fg=ModernColors.DARK_GRAY,
            relief='flat',
            padx=20,
            pady=8,
            cursor='hand2'
        )
        self.cancel_btn.pack(side='left', padx=(0, 10))
        
        # 确定按钮 - 使用标准tkinter按钮确保显示  
        self.ok_btn = tk.Button(
            button_frame,
            text="✅ 确定",
            command=self.ok,
            font=ModernFonts.BUTTON,
            bg=ModernColors.SUCCESS,
            fg=ModernColors.WHITE,
            relief='flat',
            padx=20,
            pady=8,
            cursor='hand2'
        )
        self.ok_btn.pack(side='right', padx=(10, 0))
        
        # 绑定输入变化事件
        self.bind_input_events()
        
        # 如果是分盒模版，禁用盒/小箱比例输入框
        if getattr(self, 'is_division_box', False):
            self._disable_division_box_ratio_input()
        
        # 强制初始预览更新
        self.dialog.after(100, self.update_preview)
        
    def _disable_division_box_ratio_input(self):
        """禁用分盒模版的盒/小箱比例输入框"""
        try:
            # 找到所有Entry组件
            for widget in self.dialog.winfo_children():
                self._find_and_disable_entry(widget, 'box_per_inner_case')
        except Exception as e:
            print(f"禁用输入框失败: {e}")
    
    def _find_and_disable_entry(self, widget, target_var_name):
        """递归查找并禁用特定的输入框"""
        try:
            # 检查当前组件的子组件
            for child in widget.winfo_children():
                # 如果是Entry组件，检查其关联的变量名
                if isinstance(child, tk.Entry):
                    # 检查是否是目标变量对应的Entry
                    entry_var = child.cget('textvariable')
                    if hasattr(self, f"{target_var_name}_var"):
                        target_var = getattr(self, f"{target_var_name}_var")
                        if str(entry_var) == str(target_var):
                            # 找到了目标Entry，禁用它
                            child.config(state='disabled', bg=ModernColors.LIGHT_GRAY)
                            print(f"成功禁用 {target_var_name} 输入框")
                            return
                
                # 递归搜索子组件
                self._find_and_disable_entry(child, target_var_name)
        except Exception as e:
            pass  # 忽略错误，继续搜索
    
    def create_input_row(self, parent, label_text, var_name, tooltip, default_value=""):
        """
        创建输入行
        
        Args:
            parent: 父容器
            label_text: 标签文本
            var_name: 变量名
            tooltip: 提示文本
            default_value: 默认值
        """
        # 行容器
        row_frame = tk.Frame(parent, bg=ModernColors.CARD_BG)
        row_frame.pack(fill='x', pady=8)
        
        # 标签
        label = tk.Label(
            row_frame,
            text=label_text,
            font=ModernFonts.BODY,
            bg=ModernColors.CARD_BG,
            fg=ModernColors.BLACK,
            width=15,
            anchor='w'
        )
        label.pack(side='left', padx=(0, 10))
        
        # 输入框
        entry_var = tk.StringVar(value=default_value)
        setattr(self, f"{var_name}_var", entry_var)
        
        entry = tk.Entry(
            row_frame,
            textvariable=entry_var,
            font=ModernFonts.BODY,
            width=10,
            justify='center'
        )
        entry.pack(side='left', padx=(0, 10))
        
        # 提示文本
        tip_label = tk.Label(
            row_frame,
            text=tooltip,
            font=ModernFonts.BODY_SMALL,
            bg=ModernColors.CARD_BG,
            fg=ModernColors.GRAY
        )
        tip_label.pack(side='left', fill='x', expand=True)
        
    def bind_input_events(self):
        """绑定输入变化事件"""
        # 动态绑定所有参数变量的事件
        template_config = self.get_template_config()
        for _, field_name, _, _ in template_config["params"]:
            var_name = f"{field_name}_var"
            if hasattr(self, var_name):
                var = getattr(self, var_name)
                try:
                    # 尝试使用新的trace_add方法
                    var.trace_add('write', self.update_preview)
                except AttributeError:
                    # 如果trace_add不可用，回退到旧方法
                    var.trace('w', self.update_preview)
        
        # 初始预览更新
        self.update_preview()
        
    def update_preview(self, *args):
        """更新生成预览"""
        try:
            # 动态获取所有参数值
            template_config = self.get_template_config()
            params = {}
            
            for label, field_name, _, default_value in template_config["params"]:
                var_name = f"{field_name}_var"
                if hasattr(self, var_name):
                    var = getattr(self, var_name)
                    value_str = var.get().strip()
                    value = int(value_str) if value_str else int(default_value)
                    params[field_name] = value
            
            # 假设总张数为100进行预览计算
            total_quantity = 100
            
            # 计算标签数量
            import math
            min_box_count = params.get('min_box_count', 10)
            box_count = math.ceil(total_quantity / min_box_count)
            
            if self.template_type == "set_box":
                # 套盒模版计算 - 重新理解逻辑
                min_set_count = params.get('min_set_count', 30)  # 每套张数
                boxes_per_set = params.get('boxes_per_set', 3)   # 每套盒数
                boxes_per_inner = params.get('boxes_per_inner_case', 6)  
                sets_per_outer = params.get('sets_per_outer_case', 2)
                
                # 核心计算：套数 = 总张数 ÷ 每套张数
                set_count = math.ceil(total_quantity / min_set_count)
                # 盒标数量 = 套数 × 每套盒数  
                box_count = set_count * boxes_per_set
                # 小箱数量 = 盒数 ÷ 每小箱盒数
                inner_case_count = math.ceil(box_count / boxes_per_inner)
                # 大箱数量 = 套数 ÷ 每大箱套数
                outer_case_count = math.ceil(set_count / sets_per_outer)
                
                # 套盒预览文本
                preview_text = f"""📊 生成预览 (假设总数: {total_quantity}张)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📋 当前参数设置:
• 每套张数: {min_set_count}
• 几盒为一套: {boxes_per_set}  
• 几盒入一小箱: {boxes_per_inner}
• 几套入一大箱: {sets_per_outer}

📦 套数: {set_count} 套
    每套包装: {min_set_count} 张

🗂️  盒标数量: {box_count} 个
    每套包含: {boxes_per_set} 盒

📦 小箱标数量: {inner_case_count} 个  
    每小箱包装: {boxes_per_inner} 盒

📋 大箱标数量: {outer_case_count} 个
    每大箱包装: {sets_per_outer} 套

将生成 3 个PDF文件和1个文件夹
"""
            else:
                # 常规模版计算
                box_per_inner = params.get('box_per_inner_case', 5)
                inner_per_outer = params.get('inner_case_per_outer_case', 4)
                
                inner_case_count = math.ceil(box_count / box_per_inner)
                outer_case_count = math.ceil(inner_case_count / inner_per_outer)
                
                # 常规预览文本
                preview_text = f"""📊 生成预览 (假设总数: {total_quantity}张)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📋 当前参数设置:
• 最小分盒张数: {min_box_count}
• 盒/小箱比例: {box_per_inner}  
• 小箱/大箱比例: {inner_per_outer}

🗂️  盒标数量: {box_count} 个
    每盒包装: {min_box_count} 张

📦 小箱标数量: {inner_case_count} 个  
    每箱包装: {box_per_inner} 盒

📋 大箱标数量: {outer_case_count} 个
    每箱包装: {inner_per_outer} 小箱

将生成 3 个PDF文件和1个文件夹
"""
            
            self.preview_label.config(text=preview_text, fg=ModernColors.BLACK)
            
        except ValueError as e:
            print(f"数值转换错误: {e}")
            self.preview_label.config(
                text="⚠️ 请输入有效的数字", 
                fg=ModernColors.ERROR
            )
        except Exception as e:
            print(f"预览更新失败: {e}")
            self.preview_label.config(
                text=f"❌ 预览更新失败: {str(e)}", 
                fg=ModernColors.ERROR
            )
    
    def set_config(self, config):
        """设置配置值"""
        # 参数兼容性映射 - 将旧参数名映射到新参数名
        if self.template_type == "set_box" and config:
            # 为套盒模版提供兼容性映射
            if 'box_per_inner_case' in config and 'boxes_per_inner_case' not in config:
                config['boxes_per_inner_case'] = config['box_per_inner_case']
            if 'min_box_count' in config and 'min_set_count' not in config:
                # 假设默认转换：每盒10张 → 每套30张（3盒为一套）
                config['min_set_count'] = config['min_box_count'] * 3
        elif getattr(self, 'is_division_box', False) and config:
            # 为分盒模版强制设置盒/小箱比例为1
            config['box_per_inner_case'] = 1
        
        # 动态设置所有参数
        template_config = self.get_template_config()
        for _, field_name, _, _ in template_config["params"]:
            if field_name in config:
                var_name = f"{field_name}_var"
                if hasattr(self, var_name):
                    var = getattr(self, var_name)
                    var.set(str(config[field_name]))
    
    def get_config(self):
        """获取当前配置"""
        try:
            config = {}
            template_config = self.get_template_config()
            for _, field_name, _, default_value in template_config["params"]:
                var_name = f"{field_name}_var"
                if hasattr(self, var_name):
                    var = getattr(self, var_name)
                    value = int(var.get() or default_value)
                    config[field_name] = value
            
            # 为了向后兼容，如果是套盒模版，也提供旧参数名
            if self.template_type == "set_box" and 'boxes_per_inner_case' in config:
                config['box_per_inner_case'] = config['boxes_per_inner_case']
                
            return config
        except ValueError:
            return None
    
    def validate_inputs(self):
        """验证输入"""
        try:
            config = self.get_config()
            if config is None:
                return False, "请输入有效的数字"
            
            for key, value in config.items():
                if value <= 0:
                    return False, f"{key} 必须大于0"
                # 移除1000限制，允许更大的分盒张数用于套盒模版
            
            return True, ""
        except Exception as e:
            return False, str(e)
    
    def ok(self):
        """确定按钮处理"""
        try:
            print("确定按钮被点击")
            valid, message = self.validate_inputs()
            if not valid:
                messagebox.showerror("输入错误", message)
                return
            
            self.result = self.get_config()
            print(f"保存配置: {self.result}")
            self.dialog.destroy()
        except Exception as e:
            print(f"确定按钮处理错误: {e}")
            messagebox.showerror("错误", f"保存设置时出错: {str(e)}")
    
    def cancel(self):
        """取消按钮处理"""
        try:
            print("取消按钮被点击")
            self.result = None
            self.dialog.destroy()
        except Exception as e:
            print(f"取消按钮处理错误: {e}")
            self.dialog.destroy()
    
    def show(self):
        """显示对话框并返回结果"""
        # 等待对话框关闭
        self.dialog.wait_window()
        return self.result


def show_box_label_config_dialog(parent, initial_config=None, template_type="regular", is_division_box=False):
    """
    显示盒标配置对话框的便捷函数
    
    Args:
        parent: 父窗口
        initial_config: 初始配置字典
        template_type: 模版类型 ("regular" 或 "set_box")
    
    Returns:
        dict or None: 配置结果，如果取消则返回None
    """
    dialog = BoxLabelConfigDialog(parent, initial_config, template_type, is_division_box)
    return dialog.show()