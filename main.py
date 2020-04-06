import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from tkinter import ttk
import xlrd
from xlrd.sheet import Sheet
import os
import delivery
import pivottable
import utils
from utils import Line, SheetSelector, Checkbutton
import webbrowser
import threading


window = tk.Tk()
window.title('报表处理小工具')
window.geometry('480x406')

# Func 1
tk.Label(window, text='每日线路分配').pack()

select_manifest = SheetSelector(window, '开票明细')

select_route = SheetSelector(window, '线路分配表')


def do_delivery():
    try:
        manifest = select_manifest.get()
        assert manifest, '未选择开票明细'
        route = select_route.get()
        if (not route) and does_routing.get():
            raise KeyError('未选择线路分配表，不能生成每条线路汇总')
        if (not route) and does_printer.get():
            raise KeyError('未选择线路分配表，不能生成标签打印表')
        
        does = ['routing'] if does_routing.get() else []
        does += ['daily'] if does_daily.get() else []
        does += ['printer'] if does_printer.get() else []

        f1_label.set('转换中')
        utils.log.clear()
        delivery.handle(manifest, route, does)
        f1_label.set('转换完成')
        utils.log.show()
    except Exception as e:
        raise e
        f1_label.set('')
        mb.showerror('错误', str(e))
        return


# checkbutton
does_routing = Checkbutton(window, '生成每条线路汇总').pack()
does_daily = Checkbutton(window, '生成今日汇总').pack()
does_printer = Checkbutton(window, '生成标签打印表').pack()

tk.Button(window, text="转换", command=do_delivery).pack()

f1_label = tk.StringVar()
tk.Label(window, textvariable=f1_label).pack()

ttk.Separator(window, orient='horizontal').pack(fill=tk.X)

# Func 2
tk.Label(window, text='每月明细按校分类').pack()

select_monthly = SheetSelector(window, '开票明细')


def do_pivottable():
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

    if f2_lock.locked():
        return
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
    tk.Label(l, text='项目地址：').pack(side='left')
    url = tk.Label(l, text='https://github.com/lshpku/rabbiter')
    url.pack(side='left')


def open_url(event):
    webbrowser.open('https://github.com/lshpku/rabbiter')


url.bind("<Button-1>", open_url)

window.mainloop()
