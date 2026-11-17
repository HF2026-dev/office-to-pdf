# Office转PDF v1.1

把 Excel（.xls / .xlsx）、Word（.doc / .docx）和图片
（.png / .jpg / .jpeg / .bmp / .gif / .tif / .tiff / .webp）批量转换为 PDF 的小工具。
Mac 和 Windows 都有，双击打开就能用，不联网、不上传文件。

图片转换是无损的：JPEG 原样嵌入、不重新压缩；其他格式解码后无损压缩嵌入；
透明图先合到白底上。每张图片单独生成一页 A4 的 PDF（横图用横向），按比例放进页面、不裁剪。
图片不需要 LibreOffice 就能转；只有 Excel/Word 文件需要。

## 先装 LibreOffice（免费，只装一次）

本软件调用本机安装的 LibreOffice 来转换，请先去官网免费下载安装：

https://zh-cn.libreoffice.org/download/libreoffice/

装好后打开本软件，它会自动找到 LibreOffice；找不到会在界面顶部提示。

## 使用方法

1. 点「添加文件」选一个或多个 Excel/Word/图片文件，也可以点「添加文件夹」批量添加整个文件夹里的文件。
2. 输出位置默认是各源文件所在目录；也可以取消勾选后自己选一个输出目录。
3. 点「开始转换」，转完会显示成功 / 失败数量，点「打开输出文件夹」查看 PDF。
4. 命令行也可以用：`Office转PDF 文件1.xlsx 图片1.png 文件2.docx -o 输出目录`（不加 `-o` 则输出到源文件所在目录）。

## 版本历史

- v1.1（2026-11-16）：新增图片转 PDF（无损，JPEG 原样嵌入）；图片不需要 LibreOffice 就能转；窗口标题、命令行 `--version`、安装包文件名都带版本号。
- v1.0（2026-11-12）：Excel/Word 批量转 PDF。

## 首次打开提示

- Mac：因为软件没有经过苹果公证，第一次打开时请在「访达」里右键点 `Office转PDF.app` → 选择「打开」，之后双击就能直接打开。
- Windows：第一次运行可能弹出 SmartScreen 提示，点「更多信息」→「仍要运行」即可。

## 构建方法（开发者）

GitHub Actions 自动构建：推到 `main` 分支后，`macos-13`（Intel）出 Mac 版、 `windows-latest` 出 Windows 版，成品在 Actions 的 Artifacts 里下载。

本地打包：`pip install pyinstaller`，然后

```
# Mac
pyinstaller --noconfirm --windowed --onefile --name "Office转PDF" office_to_pdf.py
# Windows
pyinstaller --noconfirm --windowed --onefile --name "Office转PDF" office_to_pdf.py
```
