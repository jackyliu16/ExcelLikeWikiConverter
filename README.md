# ExcelLikeWikiConverter

实现了一个简单、类似 excel 的表格编辑器，支持如下功能：
1. 将所有文件导出至压缩包及导入自压缩包；
2. 支持左右，上下新增行列；删除行列；
3. 支持向单元格中插入图片；阅览图片；（支持插入多个）
4. 支持将表格及其所链接的图片抓换成为 Confluence Wiki 的格式方便插入；

## Easy Packaging
```bash
$ uv add pandas xlsxwriter Pillow tksheet openpyxl pyinstaller
$ uv uv pip run pyinstaller --onefile --windowed --name "<EXPORT_FILE_NAME>" <FILE_NAME>.py
```

## Acknowledgements
本项目由 manuas, cursor 辅助生成，Gemini 辅助检查；

## Image

<img width="1920" height="1030" alt="图片" src="https://github.com/user-attachments/assets/25d38ea4-eb02-40f7-aff2-7a32e40e95be" />

Preview：https://www.bilibili.com/video/BV1qCu7zJEw4/?vd_source=a32c6b6a885f40c3040fcb76e3e5aefe
