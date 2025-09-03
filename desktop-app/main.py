#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF标签生成器 - 主程序入口
"""

import tkinter as tk
from gui.main_window import PDFLabelGeneratorApp

def main():
    """主程序入口"""
    app = PDFLabelGeneratorApp()
    app.run()

if __name__ == "__main__":
    main()