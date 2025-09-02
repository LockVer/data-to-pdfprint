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

# 添加src目录到Python路径
current_dir = Path(__file__).parent
src_dir = current_dir.parent
sys.path.insert(0, str(src_dir))

from data.excel_reader import ExcelReader
from pdf.generator import PDFGenerator

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
        self.selected_file = None
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
        
        # 生成按钮
        self.generate_btn = ModernButton(
            content,
            text="🚀  生成多级标签PDF",
            command=self.generate_pdf,
            style='success',
            width=300
        )
        self.generate_btn.pack(fill='x', pady=(0, 10))
        self.generate_btn.config(state='disabled')
        
        # 操作提示
        hint_label = tk.Label(
            content,
            text="选择Excel文件后即可生成PDF",
            font=ModernFonts.BODY_SMALL,
            bg=ModernColors.CARD_BG,
            fg=ModernColors.GRAY
        )
        hint_label.pack(fill='x')
        
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
            
            if data is not None and not data.empty:
                preview_content = "📋 数据预览 (前3行)\n"
                preview_content += "=" * 40 + "\n\n"
                
                # 显示前3行数据
                for i, (_idx, row) in enumerate(data.head(3).iterrows()):
                    preview_content += f"📌 记录 #{i+1}:\n"
                    preview_content += "-" * 25 + "\n"
                    
                    for col, value in row.items():
                        if not pd.isna(value):
                            preview_content += f"{col}: {value}\n"
                    preview_content += "\n"
                
                if len(data) > 3:
                    preview_content += f"... 还有 {len(data) - 3} 条记录\n\n"
                
                preview_content += f"✨ 总计: {len(data)} 条记录待处理"
                
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
            self.root.after(0, lambda: self.update_status("🚀 正在生成PDF...", "生成中"))
            self.root.after(0, lambda: self.progress_bar.start())
            
            # 获取数据
            data = self.excel_reader.data
            
            # 转换数据格式
            data_list = []
            for _, row in data.iterrows():
                row_data = {}
                for col, value in row.items():
                    if not pd.isna(value):
                        # 特殊处理客户编号和主题
                        if any(keyword in str(col).lower() for keyword in ['客户', 'customer', '编号', 'code']):
                            row_data['customer_code'] = str(value)
                        elif any(keyword in str(col).lower() for keyword in ['主题', 'subject', '标题']):
                            row_data['subject'] = str(value)
                        else:
                            row_data[str(col)] = str(value)
                data_list.append(row_data)
            
            # 选择输出文件
            output_file = filedialog.asksaveasfilename(
                title="保存PDF文件",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )
            
            if not output_file:
                self.root.after(0, lambda: self.progress_bar.stop())
                self.root.after(0, lambda: self.update_status("❌ 取消生成", "就绪"))
                return
            
            # 生成PDF
            self.pdf_generator.generate_multi_label_pdf(data_list, output_file)
            
            # 更新界面
            self.root.after(0, lambda: self.progress_bar.stop())
            self.root.after(0, lambda: self.update_status("🎉 PDF生成成功！", "完成"))
            self.root.after(0, lambda: messagebox.showinfo(
                "生成成功", 
                f"🎉 PDF文件已成功生成！\n\n📄 文件位置: {output_file}\n📊 包含记录: {len(data_list)} 条"
            ))
            
        except Exception as e:
            error_msg = f"生成PDF失败：{str(e)}"
            self.root.after(0, lambda: self.progress_bar.stop())
            self.root.after(0, lambda: self.update_status(f"❌ {error_msg}", "错误"))
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
    
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