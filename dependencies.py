
"""
dependencies.py

此模块负责管理应用程序的外部依赖项，包括：
- 导入必要的库，如 tkinter, tksheet, PIL (Pillow), pandas, xlsxwriter, openpyxl。
- 定义依赖项的可用性标志，以便在运行时检查。
- 提供检查依赖项的函数，并在缺少依赖项时给出提示。
"""

import tkinter as tk
from tkinter import messagebox

# 尝试导入tksheet库
# tksheet 是一个用于Tkinter的表格控件，提供了类似Excel的功能。
# 如果导入失败，则将 TKSHEET_AVAILABLE 设置为 False，并提示用户安装。
try:
    from tksheet import Sheet
    TKSHEET_AVAILABLE = True
except ImportError:
    TKSHEET_AVAILABLE = False

# 尝试导入PIL (Pillow) 库
# Pillow 是 Python 图像处理库，用于处理图片，例如从剪贴板粘贴图片。
# 如果导入失败，则将 PIL_AVAILABLE 设置为 False，并禁用相关功能。
try:
    from PIL import Image, ImageTk, ImageGrab
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# 尝试导入 pandas 库
# pandas 是一个强大的数据分析库，用于读取和处理Excel文件。
try:
    import pandas
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

# 尝试导入 xlsxwriter 库
# xlsxwriter 用于创建新的Excel .xlsx 文件。
try:
    import xlsxwriter
    XLSXWRITER_AVAILABLE = True
except ImportError:
    XLSXWRITER_AVAILABLE = False

# 尝试导入 openpyxl 库
# openpyxl 用于读取和写入 Excel .xlsx 文件。
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

def check_dependencies():
    """检查所有必要的依赖库是否已安装。"""
    missing_deps = []
    if not TKSHEET_AVAILABLE:
        missing_deps.append("tksheet")
    if not PANDAS_AVAILABLE:
        missing_deps.append("pandas")
    if not XLSXWRITER_AVAILABLE:
        missing_deps.append("xlsxwriter")
    if not OPENPYXL_AVAILABLE:
        missing_deps.append("openpyxl")

    if missing_deps:
        messagebox.showerror(
            "错误",
            f"缺少必要的依赖库：\n{', '.join(missing_deps)}\n请运行 'pip install {' '.join(missing_deps)}' 安装。"
        )
        return False
    return True

def check_pillow_availability():
    """检查 Pillow 库是否可用。"""
    if not PIL_AVAILABLE:
        print("警告：Pillow库未安装，剪贴板图片功能不可用")
        print("如需使用剪贴板功能，请运行：pip install Pillow")
    return PIL_AVAILABLE


