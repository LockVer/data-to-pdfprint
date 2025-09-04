"""
现代化美观GUI界面

采用现代设计理念的Excel转PDF工具界面
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from pathlib import Path
import sys
import pandas as pd
import math

# 添加src目录到Python路径
current_dir = Path(__file__).parent
src_dir = current_dir.parent
sys.path.insert(0, str(src_dir))

from data.excel_reader import ExcelReader
from pdf.generator import PDFGenerator
from template.box_label_template import BoxLabelTemplate
from template.inner_case_template import InnerCaseTemplate
from template.outer_case_template import OuterCaseTemplate
from template.set_box_label_template import SetBoxLabelTemplate

class ModernColors:
    """现代化配色方案"""
    
    # 主色调 - 优雅蓝绿色系
    PRIMARY = '#2E7D8A'
    PRIMARY_LIGHT = '#4A9BAC'
    PRIMARY_DARK = '#1E5A65'
    
    # 辅助色
    SECONDARY = '#F8F9FA'
    ACCENT = '#FF6B6B'
    SUCCESS = '#51CF66'
    WARNING = '#FFD93D'
    ERROR = '#FF5252'
    
    # 中性色
    WHITE = '#FFFFFF'
    LIGHT_GRAY = '#F1F3F4'
    GRAY = '#9AA0A6'
    DARK_GRAY = '#5F6368'
    BLACK = '#202124'
    
    # 卡片和阴影
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
        self.frame = tk.Frame(parent, bg=parent['bg'])
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

class StatusBadge:
    """状态徽章组件"""
    
    def __init__(self, parent, status='default'):
        self.frame = tk.Frame(parent, bg=parent['bg'])
        
        colors = {
            'success': {'bg': ModernColors.SUCCESS, 'fg': ModernColors.WHITE},
            'warning': {'bg': ModernColors.WARNING, 'fg': ModernColors.BLACK},
            'error': {'bg': ModernColors.ERROR, 'fg': ModernColors.WHITE},
            'default': {'bg': ModernColors.GRAY, 'fg': ModernColors.WHITE}
        }
        
        self.colors = colors.get(status, colors['default'])
        
        self.label = tk.Label(
            self.frame,
            font=ModernFonts.BODY_SMALL,
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            padx=12,
            pady=4,
            relief='flat'
        )
        self.label.pack()
    
    def set_text(self, text):
        self.label.config(text=text)
    
    def set_status(self, status):
        colors = {
            'success': {'bg': ModernColors.SUCCESS, 'fg': ModernColors.WHITE},
            'warning': {'bg': ModernColors.WARNING, 'fg': ModernColors.BLACK},
            'error': {'bg': ModernColors.ERROR, 'fg': ModernColors.WHITE},
            'default': {'bg': ModernColors.GRAY, 'fg': ModernColors.WHITE}
        }
        
        self.colors = colors.get(status, colors['default'])
        self.label.config(bg=self.colors['bg'], fg=self.colors['fg'])
    
    def pack(self, **kwargs):
        self.frame.pack(**kwargs)

class ModernExcelToPDFApp:
    """现代化Excel转PDF应用程序"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.excel_reader = None
        self.pdf_generator = PDFGenerator()
        self.box_label_template = BoxLabelTemplate()
        self.inner_case_template = InnerCaseTemplate()
        self.outer_case_template = OuterCaseTemplate()
        self.set_box_label_template = SetBoxLabelTemplate()
        self.selected_file = None
        self.box_label_data = None
        self.box_config = {
            'min_box_count': 10,
            'box_per_inner_case': 5,
            'inner_case_per_outer_case': 4
        }
        
        # 模板配置
        self.template_config = {
            'template_type': 'regular',  # regular, box, case
            'regular_style': 'style1'    # style1, style2
        }
        self.setup_window()
        self.create_widgets()
        
    def setup_window(self):
        """设置主窗口"""
        self.root.title("Excel转PDF工具 - 现代化版本")
        self.root.geometry("900x700")
        self.root.configure(bg=ModernColors.SECONDARY)
        self.root.resizable(True, True)
        
        # 设置最小窗口大小
        self.root.minsize(800, 600)
        
        # 窗口居中
        self.center_window()
        
        # 设置样式
        self.setup_styles()
        
    def setup_styles(self):
        """设置ttk样式"""
        style = ttk.Style()
        
        # 配置进度条样式
        style.theme_use('clam')
        style.configure(
            "Modern.Horizontal.TProgressbar",
            background=ModernColors.PRIMARY,
            troughcolor=ModernColors.LIGHT_GRAY,
            borderwidth=0,
            lightcolor=ModernColors.PRIMARY,
            darkcolor=ModernColors.PRIMARY
        )
        
    def center_window(self):
        """窗口居中显示"""
        self.root.update_idletasks()
        width = 900
        height = 700
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def create_widgets(self):
        """创建所有界面组件"""
        # 主容器
        main_container = tk.Frame(self.root, bg=ModernColors.SECONDARY)
        main_container.pack(fill='both', expand=True, padx=30, pady=30)
        
        # 标题区域
        self.create_header(main_container)
        
        # 主要内容区域
        content_frame = tk.Frame(main_container, bg=ModernColors.SECONDARY)
        content_frame.pack(fill='both', expand=True, pady=(20, 0))
        
        # 左侧面板
        left_panel = tk.Frame(content_frame, bg=ModernColors.SECONDARY, width=400)
        left_panel.pack(side='left', fill='both', expand=True, padx=(0, 15))
        left_panel.pack_propagate(False)
        
        # 右侧面板
        right_panel = tk.Frame(content_frame, bg=ModernColors.SECONDARY, width=400)
        right_panel.pack(side='right', fill='both', expand=True, padx=(15, 0))
        right_panel.pack_propagate(False)
        
        # 左侧内容
        self.create_file_selection_card(left_panel)
        self.create_file_info_card(left_panel)
        
        # 右侧内容
        self.create_control_panel(right_panel)
        self.create_data_preview_card(right_panel)
        
        # 底部状态栏
        self.create_status_bar(main_container)
        
    def create_header(self, parent):
        """创建顶部标题区域"""
        header_frame = tk.Frame(parent, bg=ModernColors.SECONDARY)
        header_frame.pack(fill='x', pady=(0, 30))
        
        # 主标题
        title_label = tk.Label(
            header_frame,
            text="📊 Excel数据转PDF工具",
            font=ModernFonts.TITLE_LARGE,
            bg=ModernColors.SECONDARY,
            fg=ModernColors.BLACK
        )
        title_label.pack(side='left')
        
        # 版本信息
        version_label = tk.Label(
            header_frame,
            text="v2.0",
            font=ModernFonts.BODY_SMALL,
            bg=ModernColors.SECONDARY,
            fg=ModernColors.GRAY
        )
        version_label.pack(side='right', pady=(10, 0))
        
    def create_file_selection_card(self, parent):
        """创建文件选择卡片"""
        card = ModernCard(parent, "文件选择")
        card.pack(fill='x', pady=(0, 20))
        content = card.get_content_frame()
        
        # 文件状态显示
        status_frame = tk.Frame(content, bg=ModernColors.CARD_BG)
        status_frame.pack(fill='x', pady=(0, 15))
        
        self.file_status_badge = StatusBadge(status_frame, 'default')
        self.file_status_badge.set_text("未选择文件")
        self.file_status_badge.pack(side='left')
        
        self.file_count_label = tk.Label(
            status_frame,
            text="数据行数: 0",
            font=ModernFonts.BODY,
            bg=ModernColors.CARD_BG,
            fg=ModernColors.GRAY
        )
        self.file_count_label.pack(side='right')
        
        # 选择文件按钮
        self.select_btn = ModernButton(
            content,
            text="📁  选择Excel文件",
            command=self.select_file,
            style='secondary'
        )
        self.select_btn.pack(fill='x', pady=(0, 10))
        
        # 拖放提示
        drop_hint = tk.Label(
            content,
            text="💡 提示: 支持 .xlsx 和 .xls 格式文件",
            font=ModernFonts.BODY_SMALL,
            bg=ModernColors.CARD_BG,
            fg=ModernColors.GRAY
        )
        drop_hint.pack(fill='x')
        
    def create_file_info_card(self, parent):
        """创建文件信息卡片"""
        card = ModernCard(parent, "文件详情")
        card.pack(fill='both', expand=True)
        content = card.get_content_frame()
        
        # 信息显示文本框
        self.info_text = scrolledtext.ScrolledText(
            content,
            bg=ModernColors.LIGHT_GRAY,
            fg=ModernColors.BLACK,
            font=ModernFonts.CODE,
            relief='flat',
            bd=0,
            wrap='word',
            padx=15,
            pady=15
        )
        self.info_text.pack(fill='both', expand=True)
        
        # 设置默认内容
        self.info_text.insert('end', "选择Excel文件后，这里将显示文件详细信息...")
        
    def create_data_preview_card(self, parent):
        """创建数据预览卡片"""
        card = ModernCard(parent, "数据预览")
        card.pack(fill='both', expand=True)
        content = card.get_content_frame()
        
        # 数据显示文本框
        self.data_text = scrolledtext.ScrolledText(
            content,
            bg=ModernColors.LIGHT_GRAY,
            fg=ModernColors.BLACK,
            font=ModernFonts.CODE,
            relief='flat',
            bd=0,
            wrap='word',
            padx=15,
            pady=15
        )
        self.data_text.pack(fill='both', expand=True)
        
        # 设置默认内容
        self.data_text.insert('end', "数据预览将在这里显示...")
        
    def create_control_panel(self, parent):
        """创建控制面板"""
        card = ModernCard(parent, "操作控制")
        card.pack(fill='x', pady=(0, 20))
        content = card.get_content_frame()
        
        # 模板选择区域
        template_frame = tk.Frame(content, bg=ModernColors.CARD_BG)
        template_frame.pack(fill='x', pady=(0, 15))
        
        # 模板标题
        template_title = tk.Label(
            template_frame,
            text="📋 模板类型",
            font=ModernFonts.BODY,
            bg=ModernColors.CARD_BG,
            fg=ModernColors.BLACK
        )
        template_title.pack(anchor='w', pady=(0, 8))
        
        # 模板类型选择
        template_type_frame = tk.Frame(template_frame, bg=ModernColors.CARD_BG)
        template_type_frame.pack(fill='x', pady=(0, 10))
        
        self.template_type_var = tk.StringVar(value='regular')
        
        # 单选按钮样式配置
        radio_style = {
            'bg': ModernColors.CARD_BG,
            'fg': ModernColors.BLACK,
            'font': ModernFonts.BODY,
            'selectcolor': ModernColors.PRIMARY_LIGHT,
            'activebackground': ModernColors.CARD_BG,
            'activeforeground': ModernColors.PRIMARY,
            'relief': 'flat',
            'highlightthickness': 0
        }
        
        tk.Radiobutton(
            template_type_frame,
            text="常规模板",
            variable=self.template_type_var,
            value='regular',
            command=self.on_template_type_change,
            **radio_style
        ).pack(side='left', padx=(0, 20))
        
        tk.Radiobutton(
            template_type_frame,
            text="分盒模板",
            variable=self.template_type_var,
            value='box',
            command=self.on_template_type_change,
            **radio_style
        ).pack(side='left', padx=(0, 20))
        
        tk.Radiobutton(
            template_type_frame,
            text="套盒模板",
            variable=self.template_type_var,
            value='case',
            command=self.on_template_type_change,
            **radio_style
        ).pack(side='left')
        
        # 常规模板外观选择（只在选择常规模板时显示）
        self.regular_style_frame = tk.Frame(template_frame, bg=ModernColors.CARD_BG)
        self.regular_style_frame.pack(fill='x', pady=(5, 0))
        
        regular_style_label = tk.Label(
            self.regular_style_frame,
            text="  🎨 外观样式:",
            font=ModernFonts.BODY_SMALL,
            bg=ModernColors.CARD_BG,
            fg=ModernColors.GRAY
        )
        regular_style_label.pack(side='left')
        
        self.regular_style_var = tk.StringVar(value='style1')
        
        tk.Radiobutton(
            self.regular_style_frame,
            text="样式1 (当前)",
            variable=self.regular_style_var,
            value='style1',
            command=self.on_regular_style_change,
            font=ModernFonts.BODY_SMALL,
            bg=ModernColors.CARD_BG,
            fg=ModernColors.GRAY,
            selectcolor=ModernColors.PRIMARY_LIGHT,
            activebackground=ModernColors.CARD_BG,
            activeforeground=ModernColors.PRIMARY,
            relief='flat',
            highlightthickness=0
        ).pack(side='left', padx=(10, 15))
        
        tk.Radiobutton(
            self.regular_style_frame,
            text="样式2 (外观2)",
            variable=self.regular_style_var,
            value='style2',
            command=self.on_regular_style_change,
            font=ModernFonts.BODY_SMALL,
            bg=ModernColors.CARD_BG,
            fg=ModernColors.GRAY,
            selectcolor=ModernColors.PRIMARY_LIGHT,
            activebackground=ModernColors.CARD_BG,
            activeforeground=ModernColors.PRIMARY,
            relief='flat',
            highlightthickness=0
        ).pack(side='left')
        
        # 分隔线
        separator = tk.Frame(content, bg=ModernColors.BORDER, height=1)
        separator.pack(fill='x', pady=(15, 15))
        
        # 进度条
        self.progress_frame = tk.Frame(content, bg=ModernColors.CARD_BG)
        self.progress_frame.pack(fill='x', pady=(0, 15))
        
        self.progress_label = tk.Label(
            self.progress_frame,
            text="就绪",
            font=ModernFonts.BODY_SMALL,
            bg=ModernColors.CARD_BG,
            fg=ModernColors.GRAY
        )
        self.progress_label.pack(anchor='w')
        
        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            style="Modern.Horizontal.TProgressbar",
            mode='indeterminate'
        )
        self.progress_bar.pack(fill='x', pady=(5, 0))
        
        # 参数配置按钮
        self.config_btn = ModernButton(
            content,
            text="⚙️  盒标参数设置",
            command=self.show_box_config,
            style='secondary',
            width=300
        )
        self.config_btn.pack(fill='x', pady=(0, 10))
        
        # 生成按钮
        self.generate_btn = ModernButton(
            content,
            text="🚀  生成盒标PDF",
            command=self.generate_pdf,
            style='success',
            width=300
        )
        self.generate_btn.pack(fill='x', pady=(0, 10))
        self.generate_btn.config(state='disabled')
        
        # 操作提示
        self.hint_label = tk.Label(
            content,
            text="选择Excel文件和模板类型后即可生成PDF",
            font=ModernFonts.BODY_SMALL,
            bg=ModernColors.CARD_BG,
            fg=ModernColors.GRAY
        )
        self.hint_label.pack(fill='x')
        
    def create_status_bar(self, parent):
        """创建状态栏"""
        self.status_frame = tk.Frame(parent, bg=ModernColors.PRIMARY, height=40)
        self.status_frame.pack(fill='x', side='bottom', pady=(20, 0))
        self.status_frame.pack_propagate(False)
        
        self.status_label = tk.Label(
            self.status_frame,
            text="✨ 欢迎使用Excel转PDF工具",
            font=ModernFonts.BODY,
            bg=ModernColors.PRIMARY,
            fg=ModernColors.WHITE
        )
        self.status_label.pack(side='left', padx=20, pady=10)
        
        # 右侧状态信息
        self.status_right = tk.Label(
            self.status_frame,
            text="准备就绪",
            font=ModernFonts.BODY_SMALL,
            bg=ModernColors.PRIMARY,
            fg=ModernColors.WHITE
        )
        self.status_right.pack(side='right', padx=20, pady=10)
        
    def select_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.selected_file = file_path
            self.load_excel_file(file_path)
            
    def load_excel_file(self, file_path):
        """加载Excel文件"""
        try:
            self.update_status("正在读取Excel文件...", "正在处理")
            self.progress_bar.start()
            
            self.excel_reader = ExcelReader(file_path)
            data = self.excel_reader.read_data()
            
            # 提取盒标特定数据
            try:
                self.box_label_data = self.excel_reader.extract_box_label_data()
            except Exception as e:
                print(f"提取盒标数据失败: {e}")
                self.box_label_data = None
            
            # 更新文件状态
            file_info = self.excel_reader.get_file_info()
            
            self.file_status_badge.set_text(f"✅ {file_info['name']}")
            self.file_status_badge.set_status('success')
            
            self.file_count_label.config(
                text=f"数据行数: {file_info['rows']}"
            )
            
            # 显示文件详细信息
            self.display_file_info(file_info)
            
            # 显示数据预览
            self.display_data_preview(data)
            
            # 启用生成按钮
            self.generate_btn.config(state='normal')
            
            self.progress_bar.stop()
            self.update_status("✅ 文件加载成功！", "就绪")
            
        except Exception as e:
            self.progress_bar.stop()
            self.file_status_badge.set_text("❌ 加载失败")
            self.file_status_badge.set_status('error')
            
            messagebox.showerror("错误", f"加载文件失败：{str(e)}")
            self.update_status(f"❌ 错误: {str(e)}", "错误")
            
    def display_file_info(self, file_info):
        """显示文件信息"""
        self.info_text.delete(1.0, 'end')
        
        info_content = f"""📁 文件信息
{'='*40}

文件名称: {file_info['name']}
文件大小: {file_info['size']} 字节 ({file_info['size_mb']} MB)
数据行数: {file_info['rows']} 行
文件状态: ✅ 读取成功

📊 数据概览
{'='*40}

数据已成功加载，可以进行PDF生成。
请查看右侧数据预览确认内容正确。
"""
        self.info_text.insert('end', info_content)
        
    def display_data_preview(self, data):
        """显示数据预览"""
        try:
            self.data_text.delete(1.0, 'end')
            
            # 优先显示盒标数据
            if self.box_label_data:
                template_type = self.template_config.get('template_type', 'regular')
                template_name = self._get_template_display_name(template_type)
                
                preview_content = f"📦 {template_name}数据预览\n"
                preview_content += "=" * 40 + "\n\n"
                preview_content += f"📋 A4 (客户名称): {self.box_label_data['A4']}\n"
                preview_content += f"🎯 B4 (主题): {self.box_label_data['B4']}\n"
                preview_content += f"🔢 B11 (开始号): {self.box_label_data['B11']}\n"
                preview_content += f"📊 F4 (总张数): {self.box_label_data['F4']}\n\n"
                
                # 计算预览信息
                total_qty = int(self.box_label_data['F4']) if str(self.box_label_data['F4']).isdigit() else 0
                if total_qty > 0:
                    box_count = math.ceil(total_qty / self.box_config.get('min_box_count', 10))
                    inner_count = math.ceil(box_count / self.box_config.get('box_per_inner_case', 5))
                    outer_count = math.ceil(inner_count / self.box_config.get('inner_case_per_outer_case', 4))
                    
                    preview_content += f"📦 {template_name}生成预览:\n"
                    
                    if template_type == 'regular':
                        style = self.template_config.get('regular_style', 'style1')
                        preview_content += f"• 模板样式: {style}\n"
                        preview_content += f"• 盒标数量: {box_count} 个\n"
                        preview_content += f"• 每盒张数: {self.box_config.get('min_box_count', 10)}\n"
                        preview_content += f"• 序号递增: 每个标签+1\n"
                        preview_content += f"• 内箱标数量: {inner_count} 个\n"
                        # 计算每箱张数，考虑参数兼容性
                        min_sheets = self.box_config.get('min_box_count', 10)
                        per_inner = self.box_config.get('box_per_inner_case', 5)
                        preview_content += f"• 每箱张数: {min_sheets * per_inner}\n\n"
                    elif template_type == 'box':
                        preview_content += f"• 分盒标数量: {box_count} 个\n"
                        preview_content += f"• 每盒张数: {self.box_config.get('min_box_count', 10)}\n"
                        preview_content += f"• 分盒配置: 待开发\n\n"
                    elif template_type == 'case':
                        preview_content += f"• 套盒标数量: {outer_count} 个\n"
                        preview_content += f"• 每套盒数: {self.box_config.get('inner_case_per_outer_case', 4)}\n"
                        preview_content += f"• 套盒配置: 待开发\n\n"
                
                if template_type == 'regular':
                    preview_content += "✅ 常规模板数据已准备就绪，可以生成PDF"
                else:
                    preview_content += f"⚠️  {template_name}正在开发中，暂时使用常规模板生成"
                    
                self.data_text.insert('end', preview_content)
                
            elif data is not None and not data.empty:
                # 显示常规数据预览
                preview_content = "📋 数据预览 (前3行)\n"
                preview_content += "=" * 40 + "\n\n"
                
                for i, (_idx, row) in enumerate(data.head(3).iterrows()):
                    preview_content += f"📌 记录 #{i+1}:\n"
                    preview_content += "-" * 25 + "\n"
                    
                    for col, value in row.items():
                        if not pd.isna(value):
                            preview_content += f"{col}: {value}\n"
                    preview_content += "\n"
                
                if len(data) > 3:
                    preview_content += f"... 还有 {len(data) - 3} 条记录\n\n"
                
                preview_content += f"⚠️  未找到盒标数据 (A4、B4、B11、F4)\n"
                preview_content += f"✨ 总计: {len(data)} 条记录"
                
                self.data_text.insert('end', preview_content)
            else:
                self.data_text.insert('end', "❌ 未找到有效数据")
                
        except Exception as e:
            self.data_text.insert('end', f"❌ 数据预览失败: {str(e)}")
    
    def generate_pdf(self):
        """生成PDF文件"""
        if not self.excel_reader:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
            
        try:
            # 在新线程中生成PDF
            threading.Thread(target=self._generate_pdf_thread, daemon=True).start()
            
        except Exception as e:
            messagebox.showerror("错误", f"生成PDF失败：{str(e)}")
            self.update_status(f"❌ 错误: {str(e)}", "错误")
    
    def _generate_pdf_thread(self):
        """在线程中生成PDF"""
        try:
            template_type = self.template_config.get('template_type', 'regular')
            template_name = self._get_template_display_name(template_type)
            
            self.root.after(0, lambda: self.update_status(f"🚀 正在生成{template_name}PDF...", "生成中"))
            self.root.after(0, lambda: self.progress_bar.start())
            
            # 检查是否有盒标数据
            if not self.box_label_data:
                raise Exception("未找到盒标数据，请确保Excel文件包含A4、B4、B11、F4位置的数据")
            
            # 选择输出目录
            output_dir = filedialog.askdirectory(
                title="选择输出目录"
            )
            
            if not output_dir:
                self.root.after(0, lambda: self.progress_bar.stop())
                self.root.after(0, lambda: self.update_status("❌ 取消生成", "就绪"))
                return
            
            # 准备数据 - 直接传递盒标数据字典
            data_dict = self.box_label_data
            
            # 输出详细信息到控制台
            print(f"开始生成{template_name}PDF...")
            print(f"输出目录: {output_dir}")
            print(f"模板配置: {self.template_config}")
            print(f"盒标数据: {data_dict}")
            print(f"配置参数: {self.box_config}")
            
            # 根据模板类型生成不同PDF
            if template_type == 'regular':
                # 常规模板 - 同时生成盒标、内箱标、外箱标
                # 获取选择的外观样式
                regular_style = self.template_config.get('regular_style', 'style1')
                appearance = 'v2' if regular_style == 'style2' else 'v1'
                
                result = self.box_label_template.generate_labels_pdf(
                    data_dict, 
                    self.box_config, 
                    output_dir,
                    appearance=appearance
                )
                
                # 生成内箱标PDF
                inner_case_result = self.inner_case_template.generate_inner_case_labels_pdf(
                    data_dict,
                    self.box_config,
                    output_dir
                )
                result.update(inner_case_result)
                
                # 生成外箱标PDF  
                outer_case_result = self.outer_case_template.generate_outer_case_labels_pdf(
                    data_dict,
                    self.box_config,
                    output_dir
                )
                result.update(outer_case_result)
            elif template_type == 'box':
                # 分盒模板（暂时使用常规模板）
                print("⚠️  分盒模板暂未实现，使用常规模板代替")
                result = self.box_label_template.generate_labels_pdf(
                    data_dict, 
                    self.box_config, 
                    output_dir
                )
            elif template_type == 'case':
                # 套盒模板 - 使用专门的套盒模版
                print("🎯 使用套盒模板生成标签")
                try:
                    # 为套盒模版适配参数格式
                    set_box_config = self.box_config.copy()
                    
                    # 确保套盒模版需要的基础参数都存在
                    set_box_config.setdefault('boxes_per_set', 3)
                    set_box_config.setdefault('boxes_per_inner_case', 6) 
                    set_box_config.setdefault('sets_per_outer_case', 2)
                    
                    # 如果没有 min_set_count，从 min_box_count 转换
                    if 'min_set_count' not in set_box_config and 'min_box_count' in set_box_config:
                        boxes_per_set = set_box_config['boxes_per_set']
                        set_box_config['min_set_count'] = set_box_config['min_box_count'] * boxes_per_set
                        print(f"🔄 参数转换: min_box_count({set_box_config['min_box_count']}) * boxes_per_set({boxes_per_set}) = min_set_count({set_box_config['min_set_count']})")
                    
                    result = self.set_box_label_template.generate_set_box_labels_pdf(
                        data_dict, 
                        set_box_config, 
                        output_dir
                    )
                except Exception as e:
                    print(f"❌ 套盒模版调用失败: {e}")
                    print(f"   回退到使用常规模版")
                    # 回退方案：使用常规模版
                    result = self.box_label_template.generate_labels_pdf(
                        data_dict, 
                        self.box_config, 
                        output_dir
                    )
            else:
                raise Exception(f"未知的模板类型: {template_type}")
            
            print(f"PDF生成结果: {result}")
            
            # 更新界面
            self.root.after(0, lambda: self.progress_bar.stop())
            self.root.after(0, lambda: self.update_status(f"🎉 {template_name}PDF生成成功！", "完成"))
            
            # 显示成功信息
            success_msg = f"""🎉 {template_name}PDF文件已成功生成！

📁 输出文件夹: {result['folder']}

生成的文件:"""

            # 根据不同模板类型显示对应的文件信息
            if 'box_labels' in result:
                success_msg += f"""
📦 盒标文件: {Path(result['box_labels']).name}"""
            elif 'set_box_labels' in result:
                success_msg += f"""
📦 套盒标文件: {Path(result['set_box_labels']).name}"""

            # 添加内箱标文件信息
            if 'inner_case_labels' in result:
                success_msg += f"""
📦 内箱标文件: {Path(result['inner_case_labels']).name}"""

            # 添加外箱标文件信息
            if 'outer_case_labels' in result:
                success_msg += f"""
📦 外箱标文件: {Path(result['outer_case_labels']).name}"""

            success_msg += f"""

📊 数据信息:
• 模板类型: {template_name}
• 客户名称 (A4): {data_dict['A4']}
• 主题 (B4): {data_dict['B4']}
• 起始编号 (B11): {data_dict['B11']}  
• 总张数 (F4): {data_dict['F4']}"""

            # 根据不同模板类型显示对应的数量信息
            if template_type == 'case':
                # 套盒模板显示套盒相关信息 - 使用result中的实际数据或计算
                if 'count' in result:
                    box_count = result['count']
                else:
                    # 备用计算：使用套盒逻辑
                    min_set_count = self.box_config.get('min_set_count', 
                                                       self.box_config.get('min_box_count', 10) * 3)
                    boxes_per_set = self.box_config.get('boxes_per_set', 3)
                    set_count = math.ceil(int(data_dict['F4']) / min_set_count)
                    box_count = set_count * boxes_per_set
                    
                boxes_per_set = self.box_config.get('boxes_per_set', 3)
                success_msg += f"""
• 套盒标数量: {box_count} 个
• 编号方式: 套号-盒号格式 (每{boxes_per_set}盒为一套)"""
            else:
                # 常规模板显示盒标信息
                success_msg += f"""
• 盒标数量: {math.ceil(int(data_dict['F4']) / self.box_config.get('min_box_count', 10))} 个
• 编号方式: 每个标签+1"""

            # 如果有内箱标信息，添加内箱标数量
            if 'inner_case_count' in result:
                success_msg += f"""
• 内箱标数量: {result['inner_case_count']} 个
• 每箱张数: {self.box_config.get('min_box_count', 10) * self.box_config.get('box_per_inner_case', 5)}"""
            
            if template_type != 'regular':
                success_msg += f"\n\n⚠️  注意: {template_name}正在开发中，当前使用常规模板生成"
            
            # 显示成功对话框，并询问是否打开文件夹
            def show_success_and_open():
                response = messagebox.askyesno(
                    "生成成功", 
                    success_msg + "\n\n是否打开输出文件夹？",
                    icon='question'
                )
                if response:
                    try:
                        import subprocess
                        import platform
                        folder_path = result['folder']
                        if platform.system() == "Darwin":  # macOS
                            subprocess.run(["open", folder_path])
                        elif platform.system() == "Windows":  # Windows
                            subprocess.run(["explorer", folder_path])
                        else:  # Linux
                            subprocess.run(["xdg-open", folder_path])
                    except Exception as e:
                        messagebox.showerror("错误", f"无法打开文件夹: {e}")
            
            self.root.after(0, show_success_and_open)
            
        except Exception as e:
            error_msg = f"生成{template_name}PDF失败：{str(e)}"
            self.root.after(0, lambda: self.progress_bar.stop())
            self.root.after(0, lambda: self.update_status(f"❌ {error_msg}", "错误"))
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
    
    def on_template_type_change(self):
        """模板类型改变时的回调"""
        template_type = self.template_type_var.get()
        self.template_config['template_type'] = template_type
        
        # 根据模板类型显示/隐藏外观选择
        if template_type == 'regular':
            self.regular_style_frame.pack(fill='x', pady=(5, 0))
        else:
            self.regular_style_frame.pack_forget()
        
        # 更新生成按钮文字和配置按钮
        if template_type == 'regular':
            self.generate_btn.button.config(text="🚀  生成盒标PDF")
            self.config_btn.button.config(text="⚙️  盒标参数设置")
        elif template_type == 'box':
            self.generate_btn.button.config(text="🚀  生成分盒PDF")  
            self.config_btn.button.config(text="⚙️  分盒参数设置")
        elif template_type == 'case':
            self.generate_btn.button.config(text="🚀  生成套盒PDF")
            self.config_btn.button.config(text="⚙️  套盒参数设置")
        
        self.update_status(f"✅ 已选择{self._get_template_display_name(template_type)}", "模板选择")
        
        # 如果已经有加载的数据，重新生成预览
        if self.box_label_data and self.excel_reader:
            data = self.excel_reader.read_data()
            self.display_data_preview(data)
    
    def on_regular_style_change(self):
        """常规模板外观改变时的回调"""
        style = self.regular_style_var.get()
        self.template_config['regular_style'] = style
        
        if style == 'style2':
            print("选择了外观2样式")
            # 外观2已经开发完成，不需要重置
        
        self.update_status(f"✅ 已选择常规模板{style}", "外观选择")
        
        # 如果已经有加载的数据，重新生成预览
        if self.box_label_data and self.excel_reader:
            data = self.excel_reader.read_data()
            self.display_data_preview(data)
    
    def _get_template_display_name(self, template_type):
        """获取模板类型的显示名称"""
        names = {
            'regular': '常规模板',
            'box': '分盒模板', 
            'case': '套盒模板'
        }
        return names.get(template_type, template_type)
    
    def show_box_config(self):
        """显示盒标参数配置对话框"""
        try:
            from .box_label_dialog import show_box_label_config_dialog
            
            # 根据当前选择的模版类型决定对话框类型
            template_type = self.template_config.get('template_type', 'regular')
            dialog_template_type = "set_box" if template_type == "case" else "regular"
            
            config = show_box_label_config_dialog(self.root, self.box_config, dialog_template_type)
            if config:
                self.box_config = config
                self.update_status("✅ 盒标参数已更新", "配置完成")
                # 如果已经有加载的数据，重新生成预览
                if self.box_label_data and self.excel_reader:
                    data = self.excel_reader.read_data()
                    self.display_data_preview(data)
        except ImportError as e:
            messagebox.showerror("错误", f"无法加载配置对话框: {e}")
        except Exception as e:
            messagebox.showerror("错误", f"配置对话框出错: {e}")
    
    def update_status(self, message, right_status=None):
        """更新状态栏"""
        self.status_label.config(text=message)
        if right_status:
            self.status_right.config(text=right_status)
            self.progress_label.config(text=right_status)
        self.root.update_idletasks()
    
    def run(self):
        """运行应用程序"""
        self.root.mainloop()

def main():
    """主函数"""
    app = ModernExcelToPDFApp()
    app.run()

if __name__ == "__main__":
    main()