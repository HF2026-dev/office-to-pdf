#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Office转PDF v1.1
把 Excel（.xls/.xlsx）、Word（.doc/.docx）和图片
（.png/.jpg/.jpeg/.bmp/.gif/.tif/.tiff/.webp）批量转换为 PDF。
Office 文件调用本机安装的免费 LibreOffice（soffice --headless）转换；
图片文件直接由本程序无损转换，不需要 LibreOffice。
图片转换无损：JPEG（RGB/灰度、无 EXIF 旋转）原样嵌入，不重新压缩；
其他格式经 Pillow 解码后用 Flate 无损压缩嵌入；透明图先合到白底上。
不联网、不上传文件。
用法：
  图形界面：直接双击运行（无参数）
  命令行：  Office转PDF 文件1.xlsx 图片1.png 文件2.docx [-o 输出目录]
            不加 -o 时输出到各源文件所在目录
"""
import argparse
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import threading
from pathlib import Path

VERSION = "1.1"
APP_NAME = "Office转PDF"
OFFICE_EXTS = {".xls", ".xlsx", ".doc", ".docx"}
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
SUPPORTED = OFFICE_EXTS | IMAGE_EXTS
LIBREOFFICE_URL = "https://zh-cn.libreoffice.org/download/libreoffice/"


def find_soffice():
    """找到本机 LibreOffice 的 soffice，找不到返回 None。"""
    p = shutil.which("soffice")
    if p:
        return p
    candidates = []
    if sys.platform == "darwin":
        candidates = ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    elif sys.platform == "win32":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
    else:
        candidates = ["/usr/bin/soffice", "/usr/local/bin/soffice",
                      "/opt/libreoffice/program/soffice"]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def convert_file(soffice, src, outdir, timeout=300):
    """转单个 Office 文件。返回 (是否成功, 说明文字)。"""
    src = Path(src)
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    target = outdir / (src.stem + ".pdf")
    profile = Path(tempfile.mkdtemp(prefix="office2pdf-profile-"))
    try:
        cmd = [str(soffice), "--headless", "--norestore",
               "-env:UserInstallation=" + profile.as_uri(),
               "--convert-to", "pdf", "--outdir", str(outdir), str(src)]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        if target.exists():
            return True, "已生成 " + target.name
        err = (r.stderr or r.stdout or "").strip()
        return False, err.splitlines()[-1][:200] if err else "转换失败，未生成 PDF"
    except subprocess.TimeoutExpired:
        return False, "转换超时（超过 %d 秒）" % timeout
    except Exception as e:  # noqa: BLE001
        return False, str(e)
    finally:
        shutil.rmtree(profile, ignore_errors=True)


def image_to_pdf(src, outdir):
    """把一张图片无损转成单页 PDF。返回 (是否成功, 说明文字)。
    JPEG（RGB/灰度、无 EXIF 旋转）原样嵌入（DCTDecode，不重新压缩）；
    其他格式经 Pillow 解码后用 Flate 无损压缩嵌入；透明图先合到白底上。
    页面为 A4（横图用横向），图片按比例放进页边距内、居中，不裁剪。
    """
    from PIL import Image, ImageOps
    import zlib

    src = Path(src)
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    target = outdir / (src.stem + ".pdf")
    try:
        im = Image.open(src)
        im.load()  # 确认文件确实能解码
    except Exception as e:  # noqa: BLE001
        return False, "打不开图片：%s" % e
    w, h = im.size
    if w <= 0 or h <= 0:
        return False, "图片尺寸无效"
    try:
        orient = im.getexif().get(0x0112, 1) if hasattr(im, "getexif") else 1
    except Exception:  # noqa: BLE001
        orient = 1

    suffix = src.suffix.lower()
    if suffix in (".jpg", ".jpeg") and orient in (1, None) and im.mode in ("RGB", "L"):
        # JPEG 原样嵌入，不重新压缩
        data = src.read_bytes()
        filt = "/DCTDecode"
        colorspace = "/DeviceRGB" if im.mode == "RGB" else "/DeviceGray"
        width, height = w, h
    else:
        # 无损通道：按 EXIF 方向转正，解码后 Flate 压缩
        img = ImageOps.exif_transpose(im)
        if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
            img = img.convert("RGBA")
            bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
            img = Image.alpha_composite(bg, img).convert("RGB")
        elif img.mode != "RGB":
            img = img.convert("RGB")
        width, height = img.size
        data = zlib.compress(img.tobytes(), 9)
        filt = "/FlateDecode"
        colorspace = "/DeviceRGB"

    # 页面：A4，横图用横向；图片按比例放进页边距内、居中
    if width >= height:
        pw, ph = 842.0, 595.0
    else:
        pw, ph = 595.0, 842.0
    margin = 36.0
    scale = min((pw - 2 * margin) / width, (ph - 2 * margin) / height)
    dw, dh = width * scale, height * scale
    x = (pw - dw) / 2
    y = (ph - dh) / 2
    content = ("q %.2f 0 0 %.2f %.2f %.2f cm /Im0 Do Q" % (dw, dh, x, y)).encode()

    objs = []
    objs.append(b"<< /Type /Catalog /Pages 2 0 R >>")
    objs.append(b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>")
    objs.append(("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 %.0f %.0f] "
                 "/Resources << /XObject << /Im0 4 0 R >> >> "
                 "/Contents 5 0 R >>" % (pw, ph)).encode())
    objs.append(("<< /Type /XObject /Subtype /Image /Width %d /Height %d "
                 "/ColorSpace %s /BitsPerComponent 8 /Filter %s "
                 "/Length %d >>" % (width, height, colorspace, filt, len(data))).encode()
                + b"\nstream\n" + data + b"\nendstream")
    objs.append(b"<< /Length %d >>\nstream\n" % len(content) + content + b"\nendstream")

    out = bytearray(b"%PDF-1.4\n")
    offsets = []
    for i, body in enumerate(objs, 1):
        offsets.append(len(out))
        out += ("%d 0 obj\n" % i).encode() + body + b"\nendobj\n"
    xref_pos = len(out)
    out += ("xref\n0 %d\n" % (len(objs) + 1)).encode()
    out += b"0000000000 65535 f \n"
    for off in offsets:
        out += ("%010d 00000 n \n" % off).encode()
    out += ("trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n"
            % (len(objs) + 1, xref_pos)).encode()
    try:
        target.write_bytes(bytes(out))
    except Exception as e:  # noqa: BLE001
        return False, "写 PDF 失败：%s" % e
    return True, "已生成 " + target.name


def convert_one(src, outdir, soffice):
    """按扩展名分流：图片直接无损转，Office 走 LibreOffice。
    soffice 为 None 时 Office 文件转不了，但图片照常转。"""
    src = Path(src)
    if src.suffix.lower() in IMAGE_EXTS:
        return image_to_pdf(src, outdir)
    if soffice is None:
        return False, "需要 LibreOffice 才能转 Office 文件（图片不需要）"
    return convert_file(soffice, src, outdir)


def convert_many(soffice, files, outdir=None, log=print):
    """批量转换。outdir 为 None 时输出到各源文件所在目录。返回 (成功数, 失败数)。"""
    ok = fail = 0
    for i, f in enumerate(files, 1):
        src = Path(f)
        dest = Path(outdir) if outdir else src.parent
        log(f"[{i}/{len(files)}] 正在转换：{src.name}")
        good, msg = convert_one(src, dest, soffice)
        if good:
            ok += 1
            log(f"    ✓ {src.name} -> {msg}")
        else:
            fail += 1
            log(f"    ✗ {src.name}：{msg}")
    return ok, fail


def run_cli(args):
    files = [a for a in args.files if Path(a).is_file()]
    skipped = [a for a in args.files if not Path(a).is_file()]
    for s in skipped:
        print(f"跳过（找不到文件）：{s}")
    files = [f for f in files if Path(f).suffix.lower() in SUPPORTED]
    if not files:
        print("没有可转换的文件。支持：.xls .xlsx .doc .docx "
              ".png .jpg .jpeg .bmp .gif .tif .tiff .webp")
        return 1
    office_files = [f for f in files if Path(f).suffix.lower() in OFFICE_EXTS]
    soffice = find_soffice() if office_files else None
    if office_files and not soffice:
        print("未找到 LibreOffice，Office 文件转不了，请先免费安装：%s" % LIBREOFFICE_URL)
        print("（图片文件不需要 LibreOffice，照常转。）")
    elif soffice:
        print(f"LibreOffice：{soffice}")
    ok, fail = convert_many(soffice, files, args.outdir)
    print(f"完成：成功 {ok} 个，失败 {fail} 个。")
    return 0 if fail == 0 else 3


# ---------------- 图形界面 ----------------

def run_gui():
    import queue
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    soffice = find_soffice()
    root = tk.Tk()
    root.title(f"{APP_NAME} v{VERSION}")
    root.geometry("680x540")
    root.minsize(620, 480)

    files_var = []
    msg_q = queue.Queue()
    busy = {"run": False}

    # --- 顶部说明 ---
    top = ttk.Frame(root, padding=10)
    top.pack(fill="x")
    ttk.Label(top, text="把 Excel / Word / 图片文件批量转换为 PDF",
              font=("", 13, "bold")).pack(anchor="w")
    ttk.Label(top, text="图片（PNG / JPG 等）直接无损转，不用 LibreOffice；Office 文件需要本机装 LibreOffice。",
              foreground="gray").pack(anchor="w", pady=(2, 0))
    if soffice:
        ttk.Label(top, text=f"已找到 LibreOffice：{soffice}",
                  foreground="green").pack(anchor="w", pady=(2, 0))
    else:
        ttk.Label(top,
                  text="未找到 LibreOffice！Office 文件转不了（图片照常转），请先免费安装：\n" + LIBREOFFICE_URL,
                  foreground="red").pack(anchor="w", pady=(2, 0))

    # --- 文件列表 ---
    mid = ttk.Frame(root, padding=(10, 0, 10, 0))
    mid.pack(fill="both", expand=True)
    listbox = tk.Listbox(mid, selectmode="extended")
    scroll = ttk.Scrollbar(mid, orient="vertical", command=listbox.yview)
    listbox.configure(yscrollcommand=scroll.set)
    listbox.pack(side="left", fill="both", expand=True)
    scroll.pack(side="right", fill="y")

    def refresh_list():
        listbox.delete(0, "end")
        for f in files_var:
            listbox.insert("end", f)

    all_patterns = "*.xls *.xlsx *.doc *.docx *.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.webp"

    def add_files():
        paths = filedialog.askopenfilenames(
            title="选择要转换的文件",
            filetypes=[("支持的文件", all_patterns),
                       ("Excel/Word 文件", "*.xls *.xlsx *.doc *.docx"),
                       ("图片文件", "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.webp"),
                       ("所有文件", "*.*")])
        for p in paths:
            if Path(p).suffix.lower() in SUPPORTED and p not in files_var:
                files_var.append(p)
            elif p:
                log(f"跳过不支持的文件：{Path(p).name}")
        refresh_list()

    def add_folder():
        d = filedialog.askdirectory(title="选择文件夹（批量转里面的文件）")
        if not d:
            return
        n = 0
        for p in sorted(Path(d).rglob("*")):
            if p.is_file() and p.suffix.lower() in SUPPORTED and str(p) not in files_var:
                files_var.append(str(p))
                n += 1
        log(f"从文件夹添加了 {n} 个文件")
        refresh_list()

    def remove_selected():
        for i in reversed(listbox.curselection()):
            files_var.pop(i)
        refresh_list()

    def clear_all():
        files_var.clear()
        refresh_list()

    btns = ttk.Frame(root, padding=(10, 6, 10, 0))
    btns.pack(fill="x")
    ttk.Button(btns, text="添加文件", command=add_files).pack(side="left", padx=(0, 6))
    ttk.Button(btns, text="添加文件夹", command=add_folder).pack(side="left", padx=(0, 6))
    ttk.Button(btns, text="移除选中", command=remove_selected).pack(side="left", padx=(0, 6))
    ttk.Button(btns, text="清空列表", command=clear_all).pack(side="left")

    # --- 输出目录 ---
    out_frame = ttk.Frame(root, padding=(10, 8, 10, 0))
    out_frame.pack(fill="x")
    same_dir = tk.BooleanVar(value=True)
    out_entry = ttk.Entry(out_frame)
    out_entry.insert(0, str(Path.home() / "Desktop"))

    def on_same_dir():
        out_entry.configure(state="disabled" if same_dir.get() else "normal")

    ttk.Checkbutton(out_frame, text="输出到各源文件所在目录",
                    variable=same_dir, command=on_same_dir).pack(anchor="w")
    row = ttk.Frame(out_frame)
    row.pack(fill="x", pady=(4, 0))
    ttk.Label(row, text="输出目录：").pack(side="left")
    out_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
    ttk.Button(row, text="选择…",
               command=lambda: (out_entry.delete(0, "end"),
                                out_entry.insert(0, filedialog.askdirectory() or out_entry.get()))
               ).pack(side="left")
    on_same_dir()

    # --- 日志 ---
    log_frame = ttk.Frame(root, padding=(10, 8, 10, 0))
    log_frame.pack(fill="both", expand=True)
    log_text = tk.Text(log_frame, height=10, state="disabled")
    log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
    log_text.configure(yscrollcommand=log_scroll.set)
    log_text.pack(side="left", fill="both", expand=True)
    log_scroll.pack(side="right", fill="y")

    def log(msg):
        msg_q.put(msg)

    def poll_log():
        while True:
            try:
                msg = msg_q.get_nowait()
            except queue.Empty:
                break
            log_text.configure(state="normal")
            log_text.insert("end", msg + "\n")
            log_text.see("end")
            log_text.configure(state="disabled")
        root.after(120, poll_log)

    # --- 底部按钮 ---
    bottom = ttk.Frame(root, padding=10)
    bottom.pack(fill="x")
    status = ttk.Label(bottom, text="就绪")
    status.pack(side="left")
    progress = ttk.Progressbar(bottom, length=180, mode="determinate")
    progress.pack(side="left", padx=10)

    def open_outdir():
        d = files_var[0] if files_var else None
        folder = Path(d).parent if (same_dir.get() and d) else Path(out_entry.get())
        try:
            if sys.platform == "darwin":
                subprocess.run(["open", str(folder)], check=False)
            elif sys.platform == "win32":
                os.startfile(str(folder))  # noqa: S606
            else:
                subprocess.run(["xdg-open", str(folder)], check=False)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("打开失败", str(e))

    def do_convert():
        if busy["run"]:
            return
        if not files_var:
            messagebox.showinfo("提示", "请先添加要转换的文件。")
            return
        office_in_list = [f for f in files_var
                          if Path(f).suffix.lower() in OFFICE_EXTS]
        if office_in_list and not soffice:
            messagebox.showerror(
                "未找到 LibreOffice",
                "列表里有 Excel/Word 文件，需要先免费安装 LibreOffice 才能转：\n"
                + LIBREOFFICE_URL + "\n（只转图片的话，把 Office 文件从列表里移除即可。）")
            return
        busy["run"] = True
        convert_btn.configure(state="disabled")
        total = len(files_var)
        progress.configure(maximum=total, value=0)
        status.configure(text="转换中…")

        def worker():
            outdir = None if same_dir.get() else out_entry.get().strip()
            ok = fail = 0
            for i, f in enumerate(list(files_var), 1):
                src = Path(f)
                dest = Path(outdir) if outdir else src.parent
                log(f"[{i}/{total}] 正在转换：{src.name}")
                good, msg = convert_one(src, dest, soffice)
                if good:
                    ok += 1
                    log(f"    ✓ {src.name} -> {msg}")
                else:
                    fail += 1
                    log(f"    ✗ {src.name}：{msg}")
                msg_q.put(("__progress__", i))
            msg_q.put(("__done__", ok, fail))

        def on_done(ok, fail):
            busy["run"] = False
            convert_btn.configure(state="normal")
            status.configure(text=f"完成：成功 {ok} 个，失败 {fail} 个")
            log(f"完成：成功 {ok} 个，失败 {fail} 个。")
            if fail:
                messagebox.showwarning("转换完成", f"成功 {ok} 个，失败 {fail} 个，详见日志。")
            else:
                messagebox.showinfo("转换完成", f"全部 {ok} 个文件已转换为 PDF。")

        def poll_done():
            while True:
                try:
                    item = msg_q.get_nowait()
                except queue.Empty:
                    break
                if isinstance(item, tuple) and item and item[0] == "__progress__":
                    progress.configure(value=item[1])
                    status.configure(text=f"转换中…（{item[1]}/{total}）")
                elif isinstance(item, tuple) and item and item[0] == "__done__":
                    on_done(item[1], item[2])
                    return
                else:
                    msg_q.put(item)
                    break
            if busy["run"]:
                root.after(150, poll_done)

        threading.Thread(target=worker, daemon=True).start()
        root.after(150, poll_done)

    convert_btn = ttk.Button(bottom, text="开始转换", command=do_convert)
    convert_btn.pack(side="right", padx=(6, 0))
    ttk.Button(bottom, text="打开输出文件夹", command=open_outdir).pack(side="right")

    log(f"{APP_NAME} v{VERSION} 已启动。支持 .xls/.xlsx/.doc/.docx 和图片 .png/.jpg/.jpeg/.bmp/.gif/.tif/.tiff/.webp 转 PDF。")
    if not soffice:
        log("未找到 LibreOffice，Office 文件转不了（图片照常转），请先安装（免费）：" + LIBREOFFICE_URL)
    root.after(120, poll_log)
    root.mainloop()


def main(argv=None):
    p = argparse.ArgumentParser(
        description="Office转PDF：把 Excel/Word/图片文件批量转换为 PDF")
    p.add_argument("files", nargs="*", help="要转换的文件（可多个）")
    p.add_argument("-o", "--outdir", default=None, help="输出目录（默认输出到各源文件所在目录）")
    p.add_argument("--version", action="version", version=f"{APP_NAME} v{VERSION}")
    args = p.parse_args(argv)
    if args.files:
        return run_cli(args)
    run_gui()
    return 0


if __name__ == "__main__":
    sys.exit(main())
