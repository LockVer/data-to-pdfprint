"""
字体管理工具类
专门负责字体注册和字体相关操作
"""

import os
import sys
import platform
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


class FontManager:
    """字体管理工具类，负责字体注册和管理"""
    
    def __init__(self):
        """初始化字体管理器"""
        self.font_name = "MicrosoftYaHei"  # 默认字体名称
        self.chinese_font_name = "MicrosoftYaHei"  # 中文字体名称
        self.font_registered = False
        
    def register_chinese_font(self):
        """
        注册中文字体
        
        Returns:
            bool: 字体注册是否成功
        """
        if self.font_registered:
            return True
            
        try:
            # 获取项目字体路径
            font_paths = self._get_font_paths()

            # 尝试注册第一个可用的字体
            for font_path in font_paths:
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(TTFont(self.font_name, font_path))
                        print(f"✅ 成功注册中文字体: {font_path}")
                        self.font_registered = True
                        return True
                    except Exception as e:
                        print(f"⚠️ 字体注册失败 {font_path}: {str(e)}")
                        continue

            # 如果没有找到合适的字体，使用Helvetica作为fallback
            print("⚠️ 未找到中文字体，将使用默认字体")
            self.font_name = "Helvetica"
            self.chinese_font_name = "Helvetica"
            return False

        except Exception as e:
            print(f"⚠️ 字体注册过程出错: {str(e)}")
            self.font_name = "Helvetica"
            self.chinese_font_name = "Helvetica"
            return False
    
    def _get_font_paths(self) -> list:
        """
        获取项目fonts目录下的字体路径列表
        考虑打包后的路径兼容性
        
        Returns:
            list: 字体路径列表
        """
        font_paths = []
        
        # 方法1: PyInstaller打包环境 - 从临时目录读取
        try:
            if getattr(sys, 'frozen', False):
                # PyInstaller打包后，字体文件在临时目录中
                base_path = sys._MEIPASS
                font_path = os.path.join(base_path, "fonts", "msyh.ttf")
                font_paths.append(font_path)
                print(f"🔍 PyInstaller模式，字体路径: {font_path}")
        except Exception as e:
            print(f"⚠️ PyInstaller路径查找失败: {e}")
        
        # 方法2: 开发环境 - 基于当前文件路径
        try:
            src_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            fonts_dir = os.path.join(src_dir, "fonts")
            font_path = os.path.join(fonts_dir, "msyh.ttf")
            font_paths.append(font_path)
            print(f"🔍 开发环境，字体路径: {font_path}")
        except Exception as e:
            print(f"⚠️ 开发环境路径查找失败: {e}")
        
        # 方法3: 相对于当前工作目录
        try:
            font_path = os.path.join("src", "fonts", "msyh.ttf")
            font_paths.append(font_path)
            print(f"🔍 相对路径，字体路径: {font_path}")
        except Exception as e:
            print(f"⚠️ 相对路径查找失败: {e}")
            
        # 去重并返回
        unique_paths = []
        for path in font_paths:
            if path not in unique_paths:
                unique_paths.append(path)
                
        print(f"🔍 所有可能的字体路径: {unique_paths}")
        return unique_paths
    
    def set_best_font(self, canvas_obj, font_size: int, bold: bool = True):
        """
        设置最适合的字体
        
        Args:
            canvas_obj: ReportLab Canvas对象
            font_size: 字体大小
            bold: 是否加粗
        """
        try:
            if bold:
                canvas_obj.setFont(self.chinese_font_name + "-Bold", font_size)
            else:
                canvas_obj.setFont(self.chinese_font_name, font_size)
        except:
            # 如果加粗字体不存在，使用常规字体
            canvas_obj.setFont(self.chinese_font_name, font_size)
    
    def has_chinese(self, text: str) -> bool:
        """
        检查文本是否包含中文字符
        
        Args:
            text: 要检查的文本
            
        Returns:
            bool: 是否包含中文字符
        """
        if not text:
            return False
        for char in text:
            if '\u4e00' <= char <= '\u9fff':
                return True
        return False
    
    def get_font_name(self) -> str:
        """获取当前字体名称"""
        return self.font_name
    
    def get_chinese_font_name(self) -> str:
        """获取中文字体名称"""
        return self.chinese_font_name
    
    def is_font_registered(self) -> bool:
        """检查字体是否已注册"""
        return self.font_registered


# 全局字体管理器实例
font_manager = FontManager()