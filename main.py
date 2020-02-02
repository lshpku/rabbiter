import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from tkinter import ttk
import xlrd
import os
import delivery
import pivottable
from utils import Line, SheetSelector
import webbrowser


window = tk.Tk()
window.title('报表处理小工具')
window.geometry('500x300')

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
        delivery.handle(manifest, route)
    except Exception as e:
        mb.showerror('错误', str(e))
        return


tk.Button(window, text="转换", command=do_delivery).pack()

ttk.Separator(window, orient='horizontal').pack(fill=tk.X)

# Func 2
tk.Label(window, text='每月明细分类').pack()

select_monthly = SheetSelector(window, '开票明细')


def do_pivottable():
    try:
        monthly = select_monthly.get()
        assert monthly, '未选择开票明细'
        pivottable.handle(monthly)
    except Exception as e:
        mb.showerror('错误', str(e))
        return


tk.Button(window, text="转换", command=do_pivottable).pack()


# Author info
text = tk.Text(window)
text.pack()

text.insert(tk.INSERT, 'https://github.com/lshpku/rabbiter')

text.tag_add('link', '1.0', '1.13')
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
