import os
from threading import Thread
import queue
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, font
from tkinterdnd2 import TkinterDnD, DND_FILES
from db import convert

import ctypes
ctypes.windll.shcore.SetProcessDpiAwareness(1)


def parse_dnd_files(data: str):
    # 兼容常见格式：{C:/a b/1.txt} {C:/2.txt} 或 C:/1.txt
    s = data.strip()
    if not s:
        return []
    if s.startswith("{") and s.endswith("}"):
        s = s[1:-1]
    if "} {" in data:
        parts = s.split("} {")
    else:
        parts = [s]
    return [p.strip() for p in parts if p.strip()]


def elide_middle(text: str, max_px: int, tkfont: font.Font):
    if max_px <= 0:
        return ""
    if tkfont.measure(text) <= max_px:
        return text

    ell = "..."
    if tkfont.measure(ell) > max_px:
        return ""

    lo, hi = 0, len(text)
    best = ell
    while lo <= hi:
        mid = (lo + hi) // 2
        a = mid // 2
        b = mid - a
        cand = (text[:a] + ell + text[-b:]) if b > 0 else (text[:a] + ell)
        if tkfont.measure(cand) <= max_px:
            best = cand
            lo = mid + 1
        else:
            hi = mid - 1
    return best


class FileDropbox:
    def __init__(self, root, name):
        self.name = name
        self.path_var = tk.StringVar(value="拖入或打开文件...")
        self.full_path = None

        outer = tk.LabelFrame(root, text=name, padx=10, pady=10)
        outer.pack(fill="x", padx=12, pady=12)

        entry = tk.Entry(outer, textvariable=self.path_var, state="readonly", width=45)
        entry.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        entry.bind("<Configure>", self.refresh_display)
        self.entry = entry

        btn = tk.Button(outer, text="打开", width=10, command=self.on_click)
        btn.grid(row=0, column=2)

        outer.grid_columnconfigure(1, weight=1)
        outer.grid_rowconfigure(0, weight=1)

        # 启用拖拽
        outer.drop_target_register(DND_FILES)
        outer.dnd_bind("<<Drop>>", self.on_drop)

    def on_drop(self, event):
        paths = parse_dnd_files(event.data)
        paths = [p for p in paths if os.path.exists(p)]
        if not paths:
            messagebox.showwarning("提示", "没有识别到有效文件路径。")
            return
        self.full_path = paths[0]
        self.refresh_display()

    def on_click(self):
        path = filedialog.askopenfilename(
            initialdir='.',
            filetypes=[('Excel文件', '.xls .xlsx'), ('所有文件', '*')]
        )
        self.full_path = path or None
        self.refresh_display()

    def refresh_display(self, event=None):
        if self.full_path is None:
            return
        max_px = max(0, self.entry.winfo_width() - 10)
        entry_font = font.Font(font=self.entry["font"])
        self.path_var.set(elide_middle(self.full_path, max_px, entry_font))


HELP_TEXT = """输出说明：
湖大：引用“湖大”进行透视，只计湖大料型
中转站销量：引用“中转站销量”进行透视
费用：引用“费用”进行透视，再乘以“折合量系数”和“元”"""


class App:
    def __init__(self, root):
        self.root = root
        root.title("销量透视表小程序")
        root.geometry("640x560")

        self.sales_book = FileDropbox(root, "明细数据表")
        self.rebate_book = FileDropbox(root, "返利费用表")

        btn = tk.Button(root, text="转换", width=10, command=self.on_click)
        btn.pack(padx=12, pady=12)
        self.convert_btn = btn
        self.running = False

        label = tk.Label(root, text=HELP_TEXT, anchor="center")
        label.pack(fill="x", pady=(8, 0))
        label.bind("<Configure>", lambda e: label.config(wraplength=label.winfo_width() - 10))
        self.progress_label = label

    def update_progress(self):
        while True:
            try:
                text = self.progress_queue.get_nowait()
            except queue.Empty:
                break
            if isinstance(text, Exception):
                messagebox.showerror("错误", str(text))
                self.progress_label.config(text="转换失败！")
                text = None
            if text is None:
                self.running = False
                self.convert_btn.config(text="转换", state="normal")
                break
            self.progress_label.config(text=text)

        if self.running:
            self.root.after(50, self.update_progress)

    def on_click(self):
        if self.running:
            return
        if self.sales_book.full_path is None:
            messagebox.showerror("错误", f"未选择{self.sales_book.name}")
            return
        if self.rebate_book.full_path is None:
            messagebox.showerror("错误", f"未选择{self.rebate_book.name}")
            return

        self.convert_btn.config(text="转换中...", state="disabled")
        self.running = True
        self.progress_queue = queue.Queue()
        t = Thread(target=self.do_convert_thread, daemon=True)
        t.start()
        self.update_progress()

    def do_convert_thread(self):
        try:
            convert(
                self.sales_book.full_path,
                self.rebate_book.full_path,
                self.progress_queue.put
            )
        except Exception:
            tb = traceback.format_exc()
            self.progress_queue.put(Exception(tb))
        finally:
            self.progress_queue.put(None)


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = App(root)
    root.mainloop()
