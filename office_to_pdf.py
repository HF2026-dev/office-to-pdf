#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Office转PDF v1.0
把 Excel（.xls/.xlsx）和 Word（.doc/.docx）文件批量转换为 PDF。
转换调用本机安装的免费 LibreOffice（soffice --headless），不联网、不上传文件。
用法：
  图形界面：直接双击运行（无参数）
  命令行：  Office转PDF 文件1.xlsx 文件2.docx [-o 输出目录]
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

VERSION = "1.0"
APP_NAME = "Office转PDF"
SUPPORTED = {".xls", ".xlsx", ".doc", ".docx"}
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
    """转单个文件。返回 (是否成功, 说明文字)。"""
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


def convert_many(soffice, files, outdir=None, log=print):
    """批量转换。outdir 为 None 时输出到各源文件所在目录。返回 (成功数, 失败数)。"""
    ok = fail = 0
    for i, f in enumerate(files, 1):
        src = Path(f)
        dest = Path(outdir) if outdir else src.parent
        log(f"[{i}/{len(files)}] 正在转换：{src.name}")
        good, msg = convert_file(soffice, src, dest)
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
        print("没有可转换的 Excel/Word 文件。支持：.xls .xlsx .doc .docx")
        return 1
    soffice = find_soffice()
    if not soffice:
        print("未找到 LibreOffice，请先免费安装：%s" % LIBREOFFICE_URL)
        return 2
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
    root.geometry("680x520")
    root.minsize(620, 460)

    files_var = []
    msg_q = queue.Queue()
    busy = {"run": False}

    # --- 顶部说明 ---
    top = ttk.Frame(root, padding=10)
    top.pack(fill="x")
    ttk.Label(top, text="把 Excel / Word 文件批量转换为 PDF",
              font=("", 13, "bold")).pack(anchor="w")
    if soffice:
        ttk.Label(top, text=f"已找到 LibreOffice：{soffice}",
                  foreground="green").pack(anchor="w", pady=(2, 0))
    else:
        ttk.Label(top,
                  text="未找到 LibreOffice！请先免费安装：\n" + LIBREOFFICE_URL,
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

    def add_files():
        paths = filedialog.askopenfilenames(
            title="选择 Excel / Word 文件",
            filetypes=[("Excel/Word 文件", "*.xls *.xlsx *.doc *.docx"),
                       ("所有文件", "*.*")])
        for p in paths:
            if Path(p).suffix.lower() in SUPPORTED and p not in files_var:
                files_var.append(p)
            elif p:
                log(f"跳过不支持的文件：{Path(p).name}")
        refresh_list()

    def add_folder():
        d = filedialog.askdirectory(title="选择文件夹（批量转里面的 Excel/Word 文件）")
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
        if not soffice:
            messagebox.showerror(
                "未找到 LibreOffice",
                "请先免费安装 LibreOffice，再用本软件转换：\n" + LIBREOFFICE_URL)
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
                good, msg = convert_file(soffice, src, dest)
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

    log(f"{APP_NAME} v{VERSION} 已启动。支持 .xls/.xlsx/.doc/.docx 转 PDF。")
    if not soffice:
        log("未找到 LibreOffice，转换前请先安装（免费）：" + LIBREOFFICE_URL)
    root.after(120, poll_log)
    root.mainloop()


def main(argv=None):
    p = argparse.ArgumentParser(description="Office转PDF：把 Excel/Word 文件批量转换为 PDF")
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
