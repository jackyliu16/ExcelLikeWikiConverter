
"""
utils.py

此模块包含各种通用辅助函数，这些函数不直接依赖于Tkinter UI或特定的业务逻辑，
可以在应用程序的不同部分复用。
"""

import os
import re
import shutil
import uuid
from datetime import datetime

class Utils:
    """通用辅助函数类，提供各种静态方法。"""

    @staticmethod
    def generate_column_headers(count):
        """
        生成Excel风格的列标题 (A, B, C, ..., Z, AA, AB, ...)。

        Args:
            count (int): 需要生成的列标题数量。

        Returns:
            list: 包含Excel风格列标题的列表。
        """
        headers = []
        for i in range(count):
            header = ""
            num = i
            while True:
                # 计算当前字符的ASCII码
                header = chr(65 + (num % 26)) + header
                num = num // 26
                if num == 0:
                    break
                num -= 1 # 调整以处理AA, AB等情况
            headers.append(header)
        return headers

    @staticmethod
    def extract_image_paths(cell_value):
        """
        从单元格值中提取图片路径。
        图片路径通常以 [IMG] 或 [IMGS] 标记开头，并可能包含多个路径，以分号分隔。

        Args:
            cell_value (str): 单元格的文本内容。

        Returns:
            list: 提取到的图片路径列表。
        """
        if not cell_value:
            return []
        
        # 使用正则表达式匹配 [IMG] 和 [IMGS] 标记后的内容
        pattern = r'\[IMGS?\]\s*([^\[\]]+)'
        matches = re.findall(pattern, str(cell_value))
        
        image_paths = []
        for match in matches:
            # 分割多个路径（用分号分隔），并去除空白字符
            paths = [path.strip() for path in match.split(';') if path.strip()]
            image_paths.extend(paths)
        
        return image_paths

    @staticmethod
    def clean_text_content(cell_value):
        """
        清理单元格内容，移除图片标记，只保留纯文本。

        Args:
            cell_value (str): 单元格的文本内容。

        Returns:
            str: 移除图片标记后的纯文本内容。
        """
        if not cell_value:
            return ""
        
        # 移除图片标记及其内容
        clean_text = re.sub(r'\[IMGS?\][^\[\]]*', '', str(cell_value))
        # 移除多余的换行符并去除首尾空白
        clean_text = re.sub(r'\n+', '\n', clean_text).strip()
        return clean_text

    @staticmethod
    def format_cell_with_images(text_content, image_paths):
        """
        格式化包含图片的单元格内容，将图片路径转换为相对路径并添加图片标记。

        Args:
            text_content (str): 单元格的纯文本内容。
            image_paths (list): 图片路径列表。

        Returns:
            str: 格式化后的单元格内容，包含文本和图片标记。
        """
        if not image_paths:
            return text_content
        
        # 确保所有路径都是相对路径
        relative_paths = []
        for path in image_paths:
            if os.path.isabs(path):
                # 如果是绝对路径，转换为相对于当前工作目录的相对路径
                relative_path = os.path.relpath(path, start=os.getcwd())
            else:
                relative_path = path
            relative_paths.append(relative_path)
        
        # 根据图片数量生成不同的图片标记
        if len(relative_paths) == 1:
            image_tag = f"[IMG] {relative_paths[0]}"
        else:
            image_tag = f"[IMGS] {'; '.join(relative_paths)}"
        
        # 将文本内容和图片标记组合起来
        if text_content and text_content.strip():
            return f"{text_content.strip()}\n{image_tag}"
        else:
            return image_tag

    @staticmethod
    def copy_images_to_assets(image_paths, assets_dir):
        """
        复制图片到assets文件夹并返回相对路径。
        处理文件名冲突，确保复制的图片文件名唯一。

        Args:
            image_paths (list): 原始图片文件路径列表。
            assets_dir (str): 目标assets目录路径。

        Returns:
            list: 复制后图片在assets目录中的相对路径列表。
        """
        relative_paths = []
        for image_path in image_paths:
            if os.path.exists(image_path):
                base_name = os.path.basename(image_path)
                name, ext = os.path.splitext(base_name)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                dest_name = f"{name}_{timestamp}{ext}"
                dest_path = os.path.join(assets_dir, dest_name)
                
                shutil.copy(image_path, dest_path)
                # 返回相对于当前工作目录的相对路径
                relative_paths.append(os.path.relpath(dest_path, start=os.getcwd()))
        
        return relative_paths

    @staticmethod
    def add_images_to_cell_incremental(sheet, row, col, new_image_paths, assets_dir):
        """
        增量添加图片到单元格（不覆盖现有内容）。
        此方法会获取单元格现有内容，提取文本和图片，然后将新图片路径添加到现有图片中，最后更新单元格内容。

        Args:
            sheet: tksheet 表格实例。
            row (int): 目标单元格的行索引。
            col (int): 目标单元格的列索引。
            new_image_paths (list): 要添加的新图片路径列表。
            assets_dir (str): 存储图片等资产的目录路径。
        """
        if not new_image_paths:
            return
        
        # 获取当前单元格内容，如果为空则初始化为""
        current_value = sheet.get_cell_data(row, col) or ""
        
        # 提取现有的文本内容和图片路径
        text_content = Utils.clean_text_content(current_value)
        existing_images = Utils.extract_image_paths(current_value)
        
        # 合并图片路径（去重但保持顺序）
        all_images = existing_images + new_image_paths
        unique_images = []
        for img in all_images:
            if img not in unique_images:
                unique_images.append(img)
        
        # 更新单元格内容
        new_value = Utils.format_cell_with_images(text_content, unique_images)
        sheet.set_cell_data(row, col, new_value)

    @staticmethod
    def auto_adjust_row_heights(sheet):
        """
        自动调整tksheet表格的行高，以适应单元格内容（包括多行文本）。

        Args:
            sheet: tksheet 表格实例。
        """
        for row in range(sheet.get_total_rows()):
            max_lines = 1
            for col in range(sheet.get_total_columns()):
                cell_value = sheet.get_cell_data(row, col)
                if cell_value:
                    # 计算单元格内容包含的行数
                    lines = str(cell_value).count('\n') + 1
                    max_lines = max(max_lines, lines)
            
            # 设置行高（每行约20像素，最小25像素）
            height = max(25, max_lines * 20)
            sheet.row_height(row=row, height=height)


