import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from tkinter import ttk
import xlrd
from xlrd.sheet import Sheet
import os
import mailtable
import utils
from utils import Line, SheetSelector
import webbrowser
import threading


window = tk.Tk()
window.title('办公室报表工具')
window.geometry('480x320')

# file selector
select_manifest = SheetSelector(window, '开票明细')
select_clilist = SheetSelector(window, '客户分配表')


def do_delivery():
    d_s = does_sales.get()
    d_a = does_annually.get()
    d_m = does_monthly.get()
    try:
        manifest = select_manifest.get()
        assert manifest, '未选择开票明细'
        clilist = select_clilist.get()
        assert not (d_s and not clilist), (
            '若要生成业务员统计，必须选择客户分配表')
        assert not (d_s and not clilist), (
            '若要生成年度统计（按客户），必须选择客户分配表')
        assert d_s or d_a or d_m, '至少选择生成一项'

        f1_label.set('转换中')
        utils.log.clear()
        mailtable.handle(manifest, clilist, d_s, d_a, d_m)
        f1_label.set('转换完成')
        utils.log.show()
    except Exception as e:
        f1_label.set('')
        mb.showerror('错误', str(e))
        return


# checkbutton
does_sales = tk.BooleanVar(value=True)
tk.Checkbutton(window, text='生成业务员统计', variable=does_sales).pack()

does_annually = tk.BooleanVar(value=True)
tk.Checkbutton(window, text='生成年度统计（按料型）', variable=does_annually).pack()

does_monthly = tk.BooleanVar(value=True)
tk.Checkbutton(window, text='生成年度统计（按客户）', variable=does_monthly).pack()

tk.Button(window, text="转换", command=do_delivery).pack()

f1_label = tk.StringVar()
tk.Label(window, textvariable=f1_label).pack()

ttk.Separator(window, orient='horizontal').pack(fill=tk.X)

# author info
with Line(window) as l:
    tk.Label(l, text='项目地址：').pack(side='left')
    url = tk.Label(l, text='https://github.com/lshpku/rabbiter')
    url.pack(side='left')


def open_url(event):
    webbrowser.open('https://github.com/lshpku/rabbiter')


url.bind("<Button-1>", open_url)

window.mainloop()
