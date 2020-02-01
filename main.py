import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from tkinter import ttk
import xlrd
import os
import delivery
import pivottable
from utils import Line, Combobox
import webbrowser

# workbooks to be opened
workbooks = {
    'manifest': None,
    'route': None,
    'monthly': None
}

window = tk.Tk()
window.title('报表处理小工具')
window.geometry('500x300')


def open_excel(name: str, label: tk.StringVar, selector: Combobox):
    filename = fd.askopenfilename(
        initialdir='.',
        filetypes=[('Excel文件', '.xls .xlsx'), ('所有文件', '*')]
    )
    if not filename:
        return None
    try:
        workbook = xlrd.open_workbook(filename)
        workbooks[name] = workbook
    except:
        mb.showerror('错误', '无法打开“{}”'.format(filename))
        return
    label.set(os.path.basename(filename))
    selector.set(workbook.sheet_names())


def open_manifest():
    open_excel('manifest', label_manifest, select_manifest)


def open_route():
    open_excel('route', label_route, select_route)


def open_monthly():
    open_excel('monthly', label_monthly, select_monthly)


tk.Label(window, text='每日线路分配').pack()

# for manifest
with Line(window) as l:
    label_manifest = tk.StringVar(value='请选择开票明细')
    tk.Label(l, textvariable=label_manifest).pack(side='left')
    tk.Button(l, text="选择文件", command=open_manifest).pack(side='left')

with Line(window) as l:
    tk.Label(l, text='选择表单').pack(side='left')
    select_manifest = Combobox(l).pack(side='left')

# for route table
with Line(window) as l:
    label_route = tk.StringVar(value='请选择线路分配表')
    tk.Label(l, textvariable=label_route).pack(side='left')
    tk.Button(l, text="选择文件", command=open_route).pack(side='left')

with Line(window) as l:
    tk.Label(l, text='选择表单').pack(side='left')
    select_route = Combobox(l).pack(side='left')


def do_delivery():
    manifest, route = workbooks['manifest'], workbooks['route']
    try:
        assert manifest, '未选择开票明细'
        assert route, '未选择线路分配表'
        mani = manifest.sheet_by_name(select_manifest.get())
        rout = route.sheet_by_name(select_route.get())
        path = 'result.xls'
        delivery.distribute(mani, rout, path)
    except Exception as e:
        mb.showerror('错误', str(e))
        return


tk.Button(window, text="转换", command=do_delivery).pack()

ttk.Separator(window, orient='horizontal').pack(fill=tk.X)
tk.Label(window, text='每月明细分类').pack()

with Line(window) as l:
    label_monthly = tk.StringVar(value='请选择开票明细')
    tk.Label(l, textvariable=label_monthly).pack(side='left')
    tk.Button(l, text="选择文件", command=open_monthly).pack(side='left')

with Line(window) as l:
    tk.Label(l, text='选择表单').pack(side='left')
    select_monthly = Combobox(l).pack(side='left')


def do_pivottable():
    monthly = workbooks['monthly']
    try:
        assert monthly, '未选择开票明细'
        mani = monthly.sheet_by_name(select_monthly.get())
        pivottable.handle_sheet(mani)
    except Exception as e:
        mb.showerror('错误', str(e))
        return


tk.Button(window, text="转换", command=do_pivottable).pack()


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
