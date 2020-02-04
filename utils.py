#
# Data structures and in-memory database.
#
import xlrd
from xlrd.sheet import Sheet
import xlwt
from xlwt import Worksheet
import tkinter as tk
from tkinter import ttk
import tkinter.filedialog as fd
import tkinter.messagebox as mb
import os
import sqlite3


class RouteMap():
    '''
    Keep a ordered list of routes and a school-route dict.
    '''

    def __init__(self, route_table: Sheet):
        self.rt_idx = {}
        self.sc_idx = {}
        self.schools = {}  # school: (route, abbr)

        routes = route_table.col_values(0)
        schools = route_table.col_values(1)
        assert len(schools) == len(routes)
        abbrs = route_table.col_values(2)
        assert len(abbrs) == len(schools)

        dist_route = set()

        for i, school in enumerate(schools):
            route, abbr = routes[i], abbrs[i]
            if (not school) or (not route) or (not abbr):
                stm = '不完整的记录：第{}行：“{} {} {}”，已丢弃'
                log(stm.format(i+1, school, route, abbr))
            else:
                self.schools[school] = (route, abbr)
                self.rt_idx.setdefault(route, len(self.rt_idx))
                self.sc_idx.setdefault(school, len(self.sc_idx))

    def sort_route(self, routes: list) -> list:
        res = [(self.rt_idx[i], i) for i in routes]
        res.sort()
        return [i[1] for i in res]

    def sort_school(self, schools: list) -> list:
        res = [(self.sc_idx[i], i) for i in schools]
        res.sort()
        return [i[1] for i in res]

    @staticmethod
    def sort_kind(kinds: list) -> list:
        order = ['肉', '菜', '油料干货', '调料制品', '杂货类']
        order = {j: i for i, j in enumerate(order)}
        res = []
        for i in kinds:
            res.append((order.get(i, len(order)), i))
        res.sort()
        return [i[1] for i in res]


class AccountDatabase():
    '''
    An in-memory database of all records.

    日期, 单据编码, 客户名称, 餐类, 货品类别, 货品编号, 货品名称,
    规格, 数量, 单位, 单价, 货款, 备注, 摘要, 仓库
    '''
    ARGS = [
        ('DATE', 'REAL', 0),        # 日期
        ('SCHOOL', 'CHAR(32)', 2),  # 客户名称
        ('MEAL', 'CHAR(16)', 3),    # 餐类
        ('KIND', 'CHAR(16)', 4),    # 货品类别
        ('NAME', 'CHAR(32)', 6),    # 货品名称
        ('SPEC', 'CHAR(32)', 7),    # 规格
        ('NUMBER', 'REAL', 8),      # 数量
        ('TOTAL', 'REAL', 11),      # 货款
    ]

    def __init__(self, sheet: Sheet):
        self.conn = sqlite3.connect(':memory:')
        self.cur = self.conn.cursor()
        self.where = None

        # create table
        self.keys = sheet.row_values(0)
        statement = ['{} {}'.format(i[0], i[1]) for i in self.ARGS]
        statement = ','.join(statement)
        self.cur.execute('CREATE TABLE TEMP({});'.format(statement))
        self.conn.commit()

        # add records
        statement = ','.join([i[0] for i in self.ARGS])
        statement = 'INSERT INTO TEMP({}) VALUES({{}})'.format(statement)
        for i in range(1, sheet.nrows):
            row = sheet.row_values(i)
            assert len(row) == len(self.keys), '不完整的行：{}'.format(i+1)
            if not row[0]:  # avoid the f**king sum line
                continue
            values = ','.join([repr(row[j[2]]) for j in self.ARGS])
            self.cur.execute(statement.format(values))
        self.conn.commit()

    def set_where(self, where=None):
        '''
        Set basic `WHERE`. This will replace the previous setting.
        '''
        self.where = where

    def select(self, target='*', where=None):
        '''
        Select with additional `WHERE`.
        '''
        w1 = self.where if self.where else ''
        w2 = where if where else ''
        wheres = ' AND '.join([w1, w2]) if w1 and w2 else w1+w2
        stm = 'SELECT {} FROM TEMP{}'.format(
            target, ' WHERE {}'.format(wheres) if wheres else '')
        return self.cur.execute(stm)


class Line():
    '''
    Add a line with `with` statement.
    '''

    def __init__(self, master=None):
        self.line = tk.Frame(master)

    def __enter__(self):
        return self.line

    def __exit__(self, exc_type, exc_value, exc_trackback):
        self.line.pack()


class Combobox():
    def __init__(self, master=None):
        self.values = tk.StringVar()
        self.cbox = ttk.Combobox(master, textvariable=self.values)

    def pack(self, cnf={}, **kw):
        self.cbox.pack(cnf, **kw)
        return self

    def set(self, values: list):
        self.cbox['value'] = values
        self.cbox.current(0)

    def get(self) -> str:
        return self.cbox.get()


class SheetSelector():
    def __init__(self, master=None, value=''):
        self.workbook = None

        with Line(master) as l:
            self.label = tk.StringVar(value='请选择{}'.format(value))
            tk.Label(l, textvariable=self.label).pack(side='left')
            tk.Button(l, text="选择文件", command=self.open).pack(side='left')

        with Line(master) as l:
            tk.Label(l, text='选择表单').pack(side='left')
            self.cbox = Combobox(l).pack(side='left')

    def open(self):
        filename = fd.askopenfilename(
            initialdir='.',
            filetypes=[('Excel文件', '.xls .xlsx'), ('所有文件', '*')]
        )
        if not filename:
            return
        try:
            self.workbook = xlrd.open_workbook(filename)
        except:
            mb.showerror('错误', '无法打开“{}”'.format(filename))
            return
        self.label.set(os.path.basename(filename))
        self.cbox.set(self.workbook.sheet_names())

    def get(self) -> Sheet:
        if not self.workbook:
            return None
        return self.workbook.sheet_by_name(self.cbox.get())


class ErrorLogger():
    def __init__(self):
        self.logs = []

    def show(self):
        if not self.logs:
            return
        win = tk.Tk()
        win.title('警告')
        win.geometry('480x240')
        frame = tk.Frame(win, padx=10, pady=10)
        text = tk.Text(frame)
        for i in self.logs:
            text.insert(tk.END, '{}\n'.format(i))
        text.pack()
        frame.pack()
        win.mainloop()

    def __call__(self, info: str):
        self.logs.append(info)

    def clear(self):
        self.logs = []


log = ErrorLogger()


if __name__ == '__main__':
    log('白兔')
    log.show()
