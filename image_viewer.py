
"""
image_viewer.py

此模块定义了 ImageViewerWindow 类，用于创建一个独立的窗口来显示图片。
它支持图片的放大、缩小、重置大小以及在图片列表中的切换。
"""

import tkinter as tk
from tkinter import ttk
import os

# 从dependencies导入PIL可用性检查
from dependencies import check_pillow_availability

# 尝试导入PIL，如果失败则禁用剪贴板图片功能
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = check_pillow_availability()
except ImportError:
    PIL_AVAILABLE = False


class ImageViewerWindow:
    """图片查看器窗口"""
    def __init__(self, parent, image_paths):
        """
        初始化图片查看器窗口。

        Args:
            parent: 父Tkinter窗口。
            image_paths (list): 包含要显示图片文件路径的列表。
        """
        self.window = tk.Toplevel(parent)
        self.window.title("图片详情")
        self.window.geometry("900x700")
        self.window.transient(parent)  # 设置为父窗口的瞬态窗口
        self.window.grab_set()  # 捕获所有事件，直到此窗口关闭
        self._is_maximized = True  # 默认最大化状态
        self._normal_geometry = self.window.geometry()  # 保存窗口正常时的几何尺寸
        self.window.state("zoomed")  # 自动最大化窗口

        # 绑定双击窗口顶部最大化/还原事件
        self.window.bind("<Double-Button-1>", self._on_titlebar_double_click)
        # 绑定Esc键关闭窗口事件
        self.window.bind("<Escape>", lambda e: self.window.destroy())

        # 主框架，用于组织窗口内的所有组件
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # 图片列表区域
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(side="left", fill="y", padx=(0, 10))
        ttk.Label(list_frame, text="图片列表:").pack(anchor="w")
        self.image_listbox = tk.Listbox(list_frame, width=30)
        self.image_listbox.pack(fill="y", expand=True)

        # 图片显示区域（Canvas和滚动条）
        display_frame = ttk.Frame(main_frame)
        display_frame.pack(side="right", fill="both", expand=True)
        self.canvas = tk.Canvas(display_frame, bg="#f0f0f0")
        self.canvas.pack(side="left", fill="both", expand=True)
        # 水平滚动条
        self.h_scroll = tk.Scrollbar(display_frame, orient="horizontal", command=self.canvas.xview)
        self.h_scroll.pack(side="bottom", fill="x")
        # 垂直滚动条
        self.v_scroll = tk.Scrollbar(display_frame, orient="vertical", command=self.canvas.yview)
        self.v_scroll.pack(side="right", fill="y")
        self.canvas.configure(xscrollcommand=self.h_scroll.set, yscrollcommand=self.v_scroll.set)
        self._canvas_img = None  # 用于存储PhotoImage对象
        # 滚轮事件
        self.canvas.bind("<MouseWheel>", self._on_mouse_wheel) # Windows/macOS
        self.canvas.bind("<Button-4>", self._on_mouse_wheel) # Linux
        self.canvas.bind("<Button-5>", self._on_mouse_wheel) # Linux

        # 按钮悬浮框架，放置在Canvas的左上角
        self.btn_frame = ttk.Frame(self.canvas, style="BtnFrame.TFrame")
        self.btn_frame.place(in_=self.canvas, x=8, y=8)
        ttk.Button(self.btn_frame, text="放大", width=4, command=self.zoom_in).pack(side="left", padx=(0, 2))
        ttk.Button(self.btn_frame, text="缩小", width=4, command=self.zoom_out).pack(side="left", padx=(0, 2))
        ttk.Button(self.btn_frame, text="原始", width=5, command=self.reset_zoom).pack(side="left", padx=(0, 2))
        ttk.Button(self.btn_frame, text="关闭", width=4, command=self.window.destroy).pack(side="left")

        # 内部数据状态
        self.image_paths = image_paths
        self.current_image = None  # 当前加载的PIL Image对象
        self.zoom_factor = 1.0  # 缩放比例
        self.current_image_path = None  # 当前显示图片的路径
        self._canvas_img_id = None  # Canvas上图片项的ID
        self._last_scroll = (0, 0)  # 记录上次滚动位置

        # 填充图片列表框
        for i, path in enumerate(image_paths):
            if os.path.exists(path):
                self.image_listbox.insert(tk.END, f"{i+1}. {os.path.basename(path)}")
            else:
                self.image_listbox.insert(tk.END, f"{i+1}. [不存在] {os.path.basename(path)}")
        # 绑定列表框选择事件
        self.image_listbox.bind("<<ListboxSelect>>", self.on_image_select)
        # 如果有图片，默认选中第一张并显示
        if self.image_listbox.size() > 0:
            self.image_listbox.selection_set(0)
            self.on_image_select()

        # 绑定Canvas的配置改变事件，用于保持滚动位置
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # 设置按钮悬浮框架的样式
        style = ttk.Style()
        style.configure("BtnFrame.TFrame", background="#f8f8f8")

    def _on_titlebar_double_click(self, event):
        """处理窗口标题栏双击事件，用于最大化/还原窗口。"""
        # 仅当点击位置靠近顶部时触发
        if event.y < 5:
            if not self._is_maximized:
                self._normal_geometry = self.window.geometry()
                self.window.state("zoomed")
                self._is_maximized = True
            else:
                self.window.state("normal")
                self.window.geometry(self._normal_geometry)
                self._is_maximized = False

    def _on_canvas_configure(self, event):
        """处理Canvas配置改变事件，用于在图片尺寸变化后保持滚动条位置。"""
        if self._last_scroll != (0, 0):
            self.canvas.xview_moveto(self._last_scroll[0])
            self.canvas.yview_moveto(self._last_scroll[1])

    def _on_mouse_wheel(self, event):
        """处理鼠标滚轮事件，用于滚动Canvas。"""
        if event.delta: # Windows/macOS
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        else: # Linux
            if event.num == 4: # 向上滚动
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5: # 向下滚动
                self.canvas.yview_scroll(1, "units")


    def on_image_select(self, event=None):
        """处理图片列表框选择事件，加载并显示选中的图片。"""
        selection = self.image_listbox.curselection()
        if not selection:
            return
        index = selection[0]
        if index < len(self.image_paths):
            self.load_image(self.image_paths[index])

    def load_image(self, image_path):
        """加载指定路径的图片文件。"""
        self.current_image_path = image_path
        if not os.path.exists(image_path):
            self.canvas.delete("all")
            self.canvas.create_text(10, 10, anchor="nw", text=f"图片不存在: {os.path.basename(image_path)}", fill="red")
            return
        if not PIL_AVAILABLE:
            self.canvas.delete("all")
            self.canvas.create_text(10, 10, anchor="nw", text=f"图片文件: {os.path.basename(image_path)}\n(需要Pillow库显示图片)", fill="black")
            return
        try:
            self.current_image = Image.open(image_path)
            self.zoom_factor = 1.0  # 每次加载新图片时重置缩放比例
            self.display_image()
        except Exception as e:
            self.canvas.delete("all")
            self.canvas.create_text(10, 10, anchor="nw", text=f"无法加载图片: {str(e)}", fill="red")

    def display_image(self):
        """在Canvas上显示当前加载的图片，并应用缩放。"""
        if not self.current_image or not PIL_AVAILABLE:
            return
        # 记录当前滚动条位置，以便在图片重新渲染后恢复
        self._last_scroll = (self.canvas.xview()[0], self.canvas.yview()[0])
        display_size = (
            int(self.current_image.width * self.zoom_factor),
            int(self.current_image.height * self.zoom_factor)
        )
        # 使用LANCZOS高质量缩放算法
        resized_image = self.current_image.resize(display_size, Image.Resampling.LANCZOS)
        self._canvas_img = ImageTk.PhotoImage(resized_image)
        self.canvas.delete("all")  # 清除Canvas上所有旧内容
        # 在Canvas上创建图片，锚点设置为西北角
        self._canvas_img_id = self.canvas.create_image(0, 0, anchor="nw", image=self._canvas_img)
        # 配置Canvas的滚动区域以适应图片大小
        self.canvas.config(scrollregion=(0, 0, display_size[0], display_size[1]))
        # 恢复滚动条位置
        self.canvas.xview_moveto(self._last_scroll[0])
        self.canvas.yview_moveto(self._last_scroll[1])

    def zoom_in(self):
        """放大图片。"""
        self.zoom_factor *= 1.2
        self.display_image()

    def zoom_out(self):
        """缩小图片。"""
        self.zoom_factor /= 1.2
        self.display_image()

    def reset_zoom(self):
        """重置图片缩放比例为原始大小。"""
        self.zoom_factor = 1.0
        self.display_image()


