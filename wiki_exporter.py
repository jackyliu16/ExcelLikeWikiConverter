
import os
from tkinter import filedialog, messagebox

from utils import Utils

class WikiExporter:
    """处理Confluence Wiki导出功能的类。"""

    def __init__(self, sheet, status_var):
        self.sheet = sheet
        self.status_var = status_var

    def export_to_wiki(self):
        """导出为Confluence Wiki文件。"""
        filename = filedialog.asksaveasfilename(
            title="导出Wiki文件",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if not filename:
            return
        try:
            wiki_content = self.get_wiki_content()
            with open(filename, "w", encoding="utf-8") as f:
                f.write(wiki_content)
            messagebox.showinfo("成功", "Wiki文件导出成功！")
            self.status_var.set("Wiki导出完成")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {str(e)}")

    def get_wiki_content(self):
        """生成Wiki内容字符串，供导出和复制复用。"""
        wiki_content = []
        # 添加表格标题
        wiki_content.append("h2. 表格数据\n\n")
        # 创建表格头
        headers = self.sheet.headers()
        header_row = "||" + "||".join(str(h) for h in headers) + "||\n"
        wiki_content.append(header_row)
        # 添加数据行
        data = self.sheet.get_sheet_data()
        for row_data in data:
            if any(cell for cell in row_data):  # 只添加非空行
                row_cells = []
                for cell_value in row_data:
                    cell_content = ""
                    if cell_value:
                        # 只保留文本内容
                        clean_text = Utils.clean_text_content(cell_value)
                        if clean_text:
                            cell_content = clean_text.replace("\n", "\\\\")
                        # 处理图片
                        image_paths = Utils.extract_image_paths(cell_value)
                        if image_paths:
                            image_refs = []
                            for image_path in image_paths:
                                if os.path.exists(image_path):
                                    image_name = os.path.basename(image_path)
                                    image_refs.append(f"!{image_name}!")
                            if image_refs:
                                if cell_content:
                                    cell_content += "|" + "|".join(image_refs)
                                else:
                                    cell_content = "|".join(image_refs)
                    row_cells.append(cell_content)
                wiki_content.append("|" + "|".join(row_cells) + "|\n")
        # 添加图片说明
        wiki_content.append("\nh3. 图片文件\n")
        wiki_content.append("请将以下图片文件上传到Confluence页面的附件中：\n\n")
        # 收集所有图片文件
        image_files = set()
        for row_data in data:
            for cell_value in row_data:
                if cell_value:
                    image_paths = Utils.extract_image_paths(cell_value)
                    for image_path in image_paths:
                        if os.path.exists(image_path):
                            image_files.add(os.path.basename(image_path))
        for image_file in sorted(image_files):
            wiki_content.append(f"* {image_file}\n")
        return "".join(wiki_content)

    def copy_wiki_to_clipboard(self):
        """复制Wiki内容到剪贴板。"""
        try:
            wiki_text = self.get_wiki_content()
            # Assuming self.root is available, which it won't be directly here.
            # This method needs to be called from SpreadsheetApp, which has access to root.
            # For now, we'll just return the text, and the app will handle clipboard.
            return wiki_text
        except Exception as e:
            messagebox.showerror("错误", f"复制Wiki到剪贴板失败: {str(e)}")
            return ""

    def _export_wiki_sync(self, filename):
        """同步导出Wiki（用于打包）。"""
        wiki_content = []
        
        # 添加表格标题
        wiki_content.append("h2. 表格数据\n\n")
        
        # 创建表格头
        headers = self.sheet.headers()
        header_row = "||" + "||".join(str(h) for h in headers) + "||\n"
        wiki_content.append(header_row)
        
        # 添加数据行
        data = self.sheet.get_sheet_data()
        for row_data in data:
            if any(cell for cell in row_data):
                row_cells = []
                for cell_value in row_data:
                    cell_content = ""
                    if cell_value:
                        clean_text = Utils.clean_text_content(cell_value)
                        if clean_text:
                            cell_content = clean_text.replace("\n", "\\\\")
                    row_cells.append(cell_content)
                
                wiki_content.append("|" + "|".join(row_cells) + "|\n")
        
        # 写入文件
        with open(filename, "w", encoding="utf-8") as f:
            f.write("".join(wiki_content))


