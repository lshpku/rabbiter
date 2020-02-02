import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from tkinter import ttk
import xlrd
from xlrd.sheet import Sheet
import os
import delivery
import pivottable
from utils import Line, SheetSelector
import webbrowser
import threading


window = tk.Tk()
window.title('报表处理小工具')
window.geometry('500x350')

# Func 1
tk.Label(window, text='每日线路分配').pack()

select_manifest = SheetSelector(window, '开票明细')

select_route = SheetSelector(window, '线路分配表')


def do_delivery():
    try:
        manifest = select_manifest.get()
        assert manifest, '未选择开票明细'
        route = select_route.get()
        assert route, '未选择线路分配表'
        f1_label.set('转换中')
        delivery.handle(manifest, route)
        f1_label.set('转换完成')
    except Exception as e:
        f1_label.set('')
        mb.showerror('错误', str(e))
        return


tk.Button(window, text="转换", command=do_delivery).pack()

f1_label = tk.StringVar()
tk.Label(window, textvariable=f1_label).pack()

ttk.Separator(window, orient='horizontal').pack(fill=tk.X)

# Func 2
tk.Label(window, text='每月明细分类').pack()

select_monthly = SheetSelector(window, '开票明细')


def do_pivottable():
    if f2_lock.locked():
        return
    def pt_thread(sheet: Sheet, log=print):
        f2_lock.acquire()
        try:
            pivottable.handle(sheet, log)
        except Exception as e:
            log('')
            f2_lock.release()
            mb.showerror('错误', str(e))
            return
        log('转换完成')
        f2_lock.release()
    try:
        monthly = select_monthly.get()
        assert monthly, '未选择开票明细'
        threading.Thread(target=pt_thread,
                         args=(monthly, f2_label.set)).start()
    except Exception as e:
        f2_label.set('')
        return


f2_lock = threading.Lock()
f2_btn = tk.StringVar(value='转换')
tk.Button(window, textvariable=f2_btn, command=do_pivottable).pack()

f2_label = tk.StringVar()
tk.Label(window, textvariable=f2_label).pack()

ttk.Separator(window, orient='horizontal').pack(fill=tk.X)


# Author info
with Line(window) as l:
    text = tk.Label(l, height=1)
    text.pack()

text.insert(tk.INSERT, '项目地址：https://github.com/lshpku/rabbiter')

text.tag_add('link', '1.5', '1.39')
text.tag_config('link', foreground='blue', underline=True)


def show_hand_cursor(event):
    text.config(cursor='arrow')


def show_xterm_cursor(event):
    text.config(cursor='xterm')


def click(event):
    webbrowser.open('https://github.com/lshpku/rabbiter')


text.tag_bind('link', '<Enter>', show_hand_cursor)
text.tag_bind('link', '<Leave>', show_xterm_cursor)
text.tag_bind('link', '<Button-1>', click)

window.mainloop()
