"""
字体工具模块

统一管理中文字体注册，供所有模版使用
优先使用微软雅黑，避免中文标点符号乱码问题
"""

import os
import platform
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def register_chinese_font():
    """
    注册中文字体 - 优先微软雅黑，避免中文标点乱码
    
    Returns:
        str: 成功注册的字体名称
    """
    try:
        system = platform.system()
        
        if system == "Darwin":  # macOS
            # 优先使用微软雅黑字体，避免中文标点乱码
            chinese_fonts = [
                # 首选微软雅黑（已找到在用户字体目录）
                (os.path.expanduser("~/Library/Fonts/Microsoft Ya Hei.ttf"), "MicrosoftYaHei"),
                ("/Library/Fonts/Microsoft YaHei.ttf", "MicrosoftYaHei"),
                ("/Library/Fonts/msyh.ttf", "MicrosoftYaHei"),
                
                # 系统自带的中文字体作为备选
                ("/System/Library/Fonts/Supplemental/Songti.ttc", "Songti"),  # 宋体
                ("/System/Library/Fonts/PingFang.ttc", "PingFang"),  # 苹方字体
                ("/System/Library/Fonts/STHeiti Medium.ttc", "STHeiti"),  # 华文黑体
                ("/System/Library/Fonts/STHeiti Light.ttc", "STHeitiLight"),  # 华文黑体细体
            ]
            
            for font_path, font_base_name in chinese_fonts:
                if os.path.exists(font_path):
                    print(f"🔍 尝试中文字体: {font_path}")
                    try:
                        if font_path.endswith('.ttc'):
                            # TTC文件尝试多个索引
                            for index in range(10):
                                try:
                                    font_name = f'{font_base_name}_{index}'
                                    pdfmetrics.registerFont(TTFont(font_name, font_path, subfontIndex=index))
                                    print(f"✅ 成功注册中文字体: {font_name}")
                                    return font_name
                                except Exception as e:
                                    continue
                        else:
                            # TTF文件直接注册
                            font_name = font_base_name
                            pdfmetrics.registerFont(TTFont(font_name, font_path))
                            print(f"✅ 成功注册中文字体: {font_name}")
                            return font_name
                    except Exception as e:
                        print(f"字体注册失败 {font_path}: {e}")
                        continue
        
        elif system == "Windows":  # Windows系统
            windows_fonts = [
                ("C:/Windows/Fonts/msyh.ttf", "MicrosoftYaHei"),  # 微软雅黑（首选）
                ("C:/Windows/Fonts/msyhbd.ttf", "MicrosoftYaHeiBold"),  # 微软雅黑粗体
                ("C:/Windows/Fonts/msyhl.ttf", "MicrosoftYaHeiLight"),  # 微软雅黑细体
                ("C:/Windows/Fonts/simhei.ttf", "SimHei"),  # 黑体
                ("C:/Windows/Fonts/simsun.ttc", "SimSun")  # 宋体
            ]
            
            for font_path, font_name in windows_fonts:
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(TTFont(font_name, font_path))
                        print(f"✅ 成功注册Windows中文字体: {font_name}")
                        return font_name
                    except Exception as e:
                        continue
        
        # 最终备用方案 - 使用系统默认字体
        print("⚠️ 未找到合适的中文字体，使用系统默认字体")
        return 'Times-Roman'  # 使用Times作为最后备选
        
    except Exception as e:
        print(f"字体注册失败: {e}")
        return 'Times-Roman'


# 全局字体缓存
_chinese_font_cache = None


def get_chinese_font():
    """
    获取已注册的中文字体名称（带缓存）
    
    Returns:
        str: 字体名称
    """
    global _chinese_font_cache
    
    if _chinese_font_cache is None:
        _chinese_font_cache = register_chinese_font()
    
    return _chinese_font_cache


def register_chinese_font_bold():
    """
    注册中文粗体字体 - 专门用于需要粗体的场合
    
    Returns:
        str: 成功注册的粗体字体名称
    """
    try:
        system = platform.system()
        
        if system == "Darwin":  # macOS
            bold_fonts = [
                # 微软雅黑粗体（如果有的话）
                (os.path.expanduser("~/Library/Fonts/Microsoft Ya Hei Bold.ttf"), "MicrosoftYaHeiBold"),
                (os.path.expanduser("~/Library/Fonts/Microsoft YaHei Bold.ttf"), "MicrosoftYaHeiBold"),
                ("/Library/Fonts/Microsoft YaHei Bold.ttf", "MicrosoftYaHeiBold"),
                
                # 使用常规微软雅黑，通过ReportLab加粗
                (os.path.expanduser("~/Library/Fonts/Microsoft Ya Hei.ttf"), "MicrosoftYaHei"),
                
                # 系统粗体字体备选
                ("/System/Library/Fonts/STHeiti Medium.ttc", "STHeiti"),
            ]
            
            for font_path, font_base_name in bold_fonts:
                if os.path.exists(font_path):
                    print(f"🔍 尝试粗体字体: {font_path}")
                    try:
                        if font_path.endswith('.ttc'):
                            # TTC文件尝试多个索引
                            for index in range(10):
                                try:
                                    font_name = f'{font_base_name}_{index}'
                                    pdfmetrics.registerFont(TTFont(font_name, font_path, subfontIndex=index))
                                    print(f"✅ 成功注册粗体字体: {font_name}")
                                    return font_name
                                except Exception as e:
                                    continue
                        else:
                            # TTF文件直接注册
                            font_name = font_base_name
                            pdfmetrics.registerFont(TTFont(font_name, font_path))
                            print(f"✅ 成功注册粗体字体: {font_name}")
                            return font_name
                    except Exception as e:
                        print(f"粗体字体注册失败 {font_path}: {e}")
                        continue
        
        elif system == "Windows":  # Windows系统
            bold_fonts = [
                ("C:/Windows/Fonts/msyhbd.ttf", "MicrosoftYaHeiBold"),  # 微软雅黑粗体
                ("C:/Windows/Fonts/msyh.ttf", "MicrosoftYaHei"),  # 微软雅黑常规
            ]
            
            for font_path, font_name in bold_fonts:
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(TTFont(font_name, font_path))
                        print(f"✅ 成功注册粗体字体: {font_name}")
                        return font_name
                    except Exception as e:
                        continue
        
        # 如果没有找到专门的粗体，返回常规字体
        print("⚠️ 未找到专门的粗体字体，使用常规字体")
        return get_chinese_font()
        
    except Exception as e:
        print(f"粗体字体注册失败: {e}")
        return get_chinese_font()


# 粗体字体缓存
_chinese_bold_font_cache = None


def get_chinese_bold_font():
    """
    获取已注册的中文粗体字体名称（带缓存）
    
    Returns:
        str: 粗体字体名称
    """
    global _chinese_bold_font_cache
    
    if _chinese_bold_font_cache is None:
        _chinese_bold_font_cache = register_chinese_font_bold()
    
    return _chinese_bold_font_cache