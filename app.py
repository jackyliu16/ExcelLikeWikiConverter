
"""
app.py

这是应用程序的主入口点，负责初始化Tkinter根窗口、创建UI组件、
以及协调各个功能模块（如文件处理、Wiki导出、图片查看等）。
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import uuid
from datetime import datetime

from tksheet import Sheet
from PIL import ImageGrab

from image_viewer import ImageViewerWindow
from file_handler import FileHandler
from wiki_exporter import WikiExporter
from utils import Utils
from dependencies import TKSHEET_AVAILABLE, PIL_AVAILABLE, check_dependencies, check_pillow_availability

class SpreadsheetApp:
    """主应用程序类，负责UI的构建和核心功能的协调。"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("表格管理软件")
        self.root.geometry("1200x800")
        
        # 数据存储
        self.assets_dir = "assets"
        # 确保assets目录存在，用于存放图片等资源
        if not os.path.exists(self.assets_dir):
            os.makedirs(self.assets_dir)
        
        # 检查依赖，如果缺少必要依赖则退出应用程序
        if not check_dependencies():
            self.root.destroy()
            return
        
        # 关键：为了解决UI布局问题，需要确保UI组件的创建顺序符合预期。
        # 工具栏应该在表格和状态栏之前创建，这样它们才能正确地pack到顶部。
        # 同时，file_handler和wiki_exporter必须在populate_toolbar之前初始化。

        # 1. 先创建工具栏框架，并将其pack到顶部
        self.create_toolbar_frame()
        
        # 2. 创建表格和状态栏，它们将pack在工具栏框架下方
        self.create_table()
        self.create_status_bar()

        # 3. 初始化功能模块实例，并将必要的依赖（如sheet、assets_dir、status_var）传递给它们
        # 这些模块的初始化依赖于self.sheet和self.status_var，所以必须在它们创建之后
        self.file_handler = FileHandler(self.sheet, self.assets_dir, self.status_var)
        self.wiki_exporter = WikiExporter(self.sheet, self.status_var)

        # 4. 填充工具栏按钮，现在file_handler和wiki_exporter已经初始化，可以安全地创建按钮
        self.populate_toolbar()
        
        # 5. 初始化数据，例如设置默认行高
        self.init_data()

    def create_toolbar_frame(self):
        """只创建工具栏框架，不包含按钮。"""
        self.toolbar = ttk.Frame(self.root)
        self.toolbar.pack(fill="x", padx=5, pady=2)

    def populate_toolbar(self):
        """填充工具栏按钮。"""
        # 文件操作按钮
        ttk.Button(self.toolbar, text="打开Excel", command=self.file_handler.open_excel_file).pack(side="left", padx=2)
        ttk.Button(self.toolbar, text="保存Excel", command=self.file_handler.save_excel_file).pack(side="left", padx=2)
        ttk.Separator(self.toolbar, orient="vertical").pack(side="left", fill="y", padx=5)
        
        # 包导入导出按钮
        ttk.Button(self.toolbar, text="导入完整包", command=self.file_handler.import_package).pack(side="left", padx=2)
        # 导出完整包时，需要将wiki_exporter实例传递给file_handler，以便file_handler可以调用wiki_exporter的同步导出方法
        ttk.Button(self.toolbar, text="导出完整包", command=lambda: self.file_handler.export_package(self.wiki_exporter)).pack(side="left", padx=2)
        ttk.Separator(self.toolbar, orient="vertical").pack(side="left", fill="y", padx=5)
        
        # 图片操作按钮
        if check_pillow_availability(): # 只有当Pillow可用时才显示粘贴图片按钮
            ttk.Button(self.toolbar, text="粘贴图片", command=self.paste_image).pack(side="left", padx=2)
        ttk.Button(self.toolbar, text="上传图片", command=self.upload_image).pack(side="left", padx=2)
        ttk.Separator(self.toolbar, orient="vertical").pack(side="left", fill="y", padx=5)
        
        # Wiki导出按钮
        ttk.Button(self.toolbar, text="导出Wiki", command=self.wiki_exporter.export_to_wiki).pack(side="left", padx=2)
        ttk.Button(self.toolbar, text="复制Wiki到剪贴板", command=self.copy_wiki_to_clipboard).pack(side="left", padx=2)
        
        # 显示Pillow库是否可用的状态提示
        if not check_pillow_availability():
            status_label = ttk.Label(self.toolbar, text="(Pillow未安装，剪贴板功能不可用)", foreground="orange")
            status_label.pack(side="right", padx=10)
    
    def create_table(self):
        """创建tksheet表格组件。"""
        table_frame = ttk.Frame(self.root)
        # 表格框架应该在工具栏下方，并填充剩余空间
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        initial_columns = 20
        # 生成Excel风格的列标题 (A, B, C, ...)
        column_headers = Utils.generate_column_headers(initial_columns)
        
        # 初始化tksheet表格
        self.sheet = Sheet(
            table_frame,
            data=[[""] * initial_columns for _ in range(100)], # 初始100行20列的空数据
            headers=column_headers,
            row_index=None, # 不使用自定义行索引，让tksheet自动管理
            width=1150,
            height=600
        )
        
        # 配置tksheet的绑定事件，启用各种交互功能
        self.sheet.enable_bindings([
            "single_select", "row_select", "column_select", "drag_select",
            "select_all", "edit_cell", "copy", "paste", "delete",
            "rc_select", "arrowkeys", "row_width_resize", "column_width_resize",
            "double_click_column_resize", "row_height_resize"
        ])
        
        # 绑定自定义事件处理函数
        self.sheet.bind("<<SheetModified>>", self.on_cell_modified) # 单元格内容修改事件
        self.sheet.bind("<Double-Button-1>", self.on_double_click) # 双击事件
        self.sheet.bind("<Button-3>", self.on_right_click) # 右键点击事件
        self.sheet.bind("<Control-MouseWheel>", self.on_ctrl_scroll) # Ctrl+滚轮事件
        
        self.sheet.pack(fill="both", expand=True)
    
    def create_status_bar(self):
        """创建状态栏，显示应用程序状态和提示信息。"""
        self.status_var = tk.StringVar()
        self.status_var.set("就绪 - 提示：双击单元格查看图片，右键查看菜单，Ctrl+滚轮调整行列宽度")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken")
        # 状态栏应该在最底部
        status_bar.pack(fill="x", side="bottom")
    
    def init_data(self):
        """初始化数据，设置默认行高。"""
        # 设置默认行高以支持多行文本显示
        for row in range(self.sheet.get_total_rows()):
            self.sheet.row_height(row=row, height=25)
    
    def update_row_index(self):
        """更新行号索引。"""
        total_rows = self.sheet.get_total_rows()
        row_index = [str(i+1) for i in range(total_rows)]
        self.sheet.row_index(row_index)
    
    def update_column_headers(self):
        """更新列标题。"""
        total_columns = self.sheet.get_total_columns()
        column_headers = Utils.generate_column_headers(total_columns)
        self.sheet.headers(column_headers)
    
    def on_cell_modified(self, event):
        """单元格修改事件处理，自动调整行高。"""
        Utils.auto_adjust_row_heights(self.sheet)
    
    def on_double_click(self, event):
        """双击单元格事件处理，用于查看图片。"""
        selected = self.sheet.get_currently_selected()
        if not selected:
            return
        
        row, col = selected.row, selected.column
        cell_value = self.sheet.get_cell_data(row, col)
        
        # 从单元格内容中提取图片路径，并打开图片查看器
        image_paths = Utils.extract_image_paths(cell_value)
        if image_paths:
            ImageViewerWindow(self.root, image_paths)
    
    def on_ctrl_scroll(self, event):
        """Ctrl+滚轮事件处理，批量调整所有行高和所有列宽。"""
        try:
            if event.delta > 0:
                scale_factor = 1.1 # 放大
            else:
                scale_factor = 0.9 # 缩小

            # 批量调整所有行高
            for row in range(self.sheet.get_total_rows()):
                try:
                    current_height = self.sheet.row_height(row=row)
                    if not isinstance(current_height, int):
                        current_height = 25 # 默认值
                except Exception:
                    current_height = 25
                new_height = int(max(21, min(200, current_height * scale_factor)))
                self.sheet.row_height(row=row, height=new_height)

            # 批量调整所有列宽
            for col in range(self.sheet.get_total_columns()):
                try:
                    current_width = self.sheet.column_width(column=col)
                    if not isinstance(current_width, int):
                        current_width = 100 # 默认值
                except Exception:
                    current_width = 100
                new_width = int(max(30, min(500, current_width * scale_factor)))
                self.sheet.column_width(column=col, width=new_width)

            self.status_var.set(f"所有行高和列宽已同步调整")
        except Exception as e:
            print(f"Ctrl+滚轮调整错误: {e}")
    
    def get_selection_type(self):
        """获取当前选择类型（行、列或单元格）。"""
        try:
            selected_rows = self.sheet.get_selected_rows(return_tuple=True)
            selected_columns = self.sheet.get_selected_columns(return_tuple=True)
            
            # 判断是否选中整行
            if selected_rows and len(selected_rows) > 0:
                if not selected_columns or len(selected_columns) == 0: # 如果有行选择但没有列选择，则认为是整行选择
                    return "row"
            
            # 判断是否选中整列
            if selected_columns and len(selected_columns) > 0:
                if not selected_rows or len(selected_rows) == 0: # 如果有列选择但没有行选择，则认为是整列选择
                    return "column"
            
            return "cell" # 否则是单元格选择
            
        except Exception as e:
            print(f"获取选择类型错误: {e}")
            return "cell"
    
    def on_right_click(self, event):
        """右键菜单事件处理，根据选择类型显示不同的上下文菜单。"""
        selection_type = self.get_selection_type()
        
        context_menu = tk.Menu(self.root, tearoff=0)
        
        # 根据选择类型添加菜单项
        if selection_type == "row":
            context_menu.add_command(label="在上方插入行", command=self.insert_row_above)
            context_menu.add_command(label="在下方插入行", command=self.insert_row_below)
            context_menu.add_command(label="删除行", command=self.delete_row)
            self.status_var.set("选中整行 - 显示行操作菜单")
        elif selection_type == "column":
            context_menu.add_command(label="在左侧插入列", command=self.insert_column_left)
            context_menu.add_command(label="在右侧插入列", command=self.insert_column_right)
            context_menu.add_command(label="删除列", command=self.delete_column)
            self.status_var.set("选中整列 - 显示列操作菜单")
        else:
            # 单元格选择时显示所有操作
            context_menu.add_command(label="在上方插入行", command=self.insert_row_above)
            context_menu.add_command(label="在下方插入行", command=self.insert_row_below)
            context_menu.add_command(label="删除行", command=self.delete_row)
            context_menu.add_separator()
            
            context_menu.add_command(label="在左侧插入列", command=self.insert_column_left)
            context_menu.add_command(label="在右侧插入列", command=self.insert_column_right)
            context_menu.add_command(label="删除列", command=self.delete_column)
            context_menu.add_separator()
            
            if check_pillow_availability():
                context_menu.add_command(label="粘贴图片", command=self.paste_image)
            context_menu.add_command(label="上传图片", command=self.upload_image)
            self.status_var.set("选中单元格 - 显示完整菜单")
        
        # 显示右键菜单
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()
    
    def paste_image(self):
        """粘贴图片到当前选中单元格。"""
        if not check_pillow_availability():
            messagebox.showerror("错误", "需要安装Pillow库来支持剪贴板图片功能\n请运行: pip install Pillow")
            return
        
        selected = self.sheet.get_currently_selected()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个单元格")
            return
        
        try:
            clipboard_data = ImageGrab.grabclipboard()
            
            if clipboard_data is None:
                return # 剪贴板中没有图片数据
            
            new_image_paths = []
            
            if isinstance(clipboard_data, list): # 如果剪贴板内容是文件路径列表
                image_files = [f for f in clipboard_data if f.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff"))]
                if image_files:
                    new_image_paths = Utils.copy_images_to_assets(image_files, self.assets_dir)
            else: # 如果剪贴板内容是PIL Image对象
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"clipboard_{timestamp}_{uuid.uuid4().hex[:8]}.png"
                filepath = os.path.join(self.assets_dir, filename)
                clipboard_data.save(filepath)
                new_image_paths = [os.path.relpath(filepath, start=os.getcwd())]
            
            if new_image_paths:
                # 增量添加图片路径到单元格内容
                Utils.add_images_to_cell_incremental(self.sheet, selected.row, selected.column, new_image_paths, self.assets_dir)
                self.status_var.set(f"已粘贴 {len(new_image_paths)} 张图片到单元格 ({selected.row+1}, {Utils.generate_column_headers(selected.column+1)[selected.column]})")
            
        except Exception as e:
            messagebox.showerror("错误", f"粘贴图片失败: {str(e)}")
    
    def upload_image(self):
        """上传图片到当前选中单元格。"""
        selected = self.sheet.get_currently_selected()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个单元格")
            return
        
        filetypes = [
            ("图片文件", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff"),
            ("所有文件", "*.*")
        ]
        
        filenames = filedialog.askopenfilenames(
            title="选择图片文件",
            filetypes=filetypes
        )
        
        if filenames:
            new_image_paths = Utils.copy_images_to_assets(filenames, self.assets_dir)
            if new_image_paths:
                # 增量添加图片路径到单元格内容
                Utils.add_images_to_cell_incremental(self.sheet, selected.row, selected.column, new_image_paths, self.assets_dir)
                self.status_var.set(f"已上传 {len(new_image_paths)} 张图片到单元格 ({selected.row+1}, {Utils.generate_column_headers(selected.column+1)[selected.column]})")
    
    def insert_row_above(self):
        """在选中行上方插入行。"""
        selected = self.sheet.get_currently_selected()
        if selected:
            self.sheet.insert_rows(rows=1, idx=selected.row)
            self.update_row_index()
            self.status_var.set(f"已在第 {selected.row+1} 行上方插入新行")
        else:
            self.sheet.insert_rows(rows=1, idx=0)
            self.update_row_index()
            self.status_var.set("已在顶部插入新行")
    
    def insert_row_below(self):
        """在选中行下方插入行。"""
        selected = self.sheet.get_currently_selected()
        if selected:
            self.sheet.insert_rows(rows=1, idx=selected.row + 1)
            self.update_row_index()
            self.status_var.set(f"已在第 {selected.row+1} 行下方插入新行")
        else:
            self.sheet.insert_rows(rows=1)
            self.update_row_index()
            self.status_var.set("已在底部插入新行")
    
    def delete_row(self):
        """删除选中行。"""
        selected = self.sheet.get_currently_selected()
        if not selected:
            messagebox.showwarning("警告", "请先选择要删除的行")
            return
        
        if self.sheet.get_total_rows() <= 1:
            messagebox.showwarning("警告", "至少需要保留一行")
            return
        
        try:
            self.sheet.delete_rows(rows=selected.row)
            self.update_row_index()
            self.status_var.set(f"已删除第 {selected.row+1} 行")
        except Exception as e:
            messagebox.showerror("错误", f"删除行失败: {str(e)}")
            print(f"删除行错误详情: {e}")
    
    def insert_column_left(self):
        """在选中列左侧插入列。"""
        selected = self.sheet.get_currently_selected()
        if selected:
            self.sheet.insert_columns(columns=1, idx=selected.column)
            self.update_column_headers()
            column_name = Utils.generate_column_headers(selected.column+1)[selected.column]
            self.status_var.set(f"已在列 {column_name} 左侧插入新列")
        else:
            self.sheet.insert_columns(columns=1, idx=0)
            self.update_column_headers()
            self.status_var.set("已在最左侧插入新列")
    
    def insert_column_right(self):
        """在选中列右侧插入列。"""
        selected = self.sheet.get_currently_selected()
        if selected:
            self.sheet.insert_columns(columns=1, idx=selected.column + 1)
            self.update_column_headers()
            column_name = Utils.generate_column_headers(selected.column+2)[selected.column+1]
            self.status_var.set(f"已在列 {Utils.generate_column_headers(selected.column+1)[selected.column]} 右侧插入新列")
        else:
            self.sheet.insert_columns(columns=1)
            self.update_column_headers()
            self.status_var.set("已在最右侧插入新列")
    
    def delete_column(self):
        """删除选中列。"""
        selected = self.sheet.get_currently_selected()
        if not selected:
            messagebox.showwarning("警告", "请先选择要删除的列")
            return
        
        if self.sheet.get_total_columns() <= 1:
            messagebox.showwarning("警告", "至少需要保留一列")
            return
        
        try:
            column_name = Utils.generate_column_headers(selected.column+1)[selected.column]
            self.sheet.delete_columns(columns=selected.column)
            self.update_column_headers()
            self.status_var.set(f"已删除列 {column_name}")
        except Exception as e:
            messagebox.showerror("错误", f"删除列失败: {str(e)}")
            print(f"删除列错误详情: {e}")

    def copy_wiki_to_clipboard(self):
        """复制Wiki内容到剪贴板。"""
        try:
            # 调用WikiExporter实例的方法获取Wiki内容
            wiki_text = self.wiki_exporter.get_wiki_content()
            self.root.clipboard_clear()
            self.root.clipboard_append(wiki_text)
            self.status_var.set("Wiki内容已复制到剪贴板")
            messagebox.showinfo("成功", "Wiki内容已复制到剪贴板！")
        except Exception as e:
            messagebox.showerror("错误", f"复制Wiki到剪贴板失败: {str(e)}")

    def run(self):
        """运行应用程序主循环。"""
        self.root.mainloop()


if __name__ == "__main__":
    # 在此处添加应用程序的启动逻辑
    # 检查所有依赖是否满足，如果满足则创建并运行应用实例
    # 依赖检查已在 SpreadsheetApp.__init__ 中完成，这里只需创建实例并运行
    app = SpreadsheetApp()
    app.run()


