#
# Data structures and in-memory database.
#
import xlrd
from xlrd.sheet import Sheet
import xlwt
from xlwt import Worksheet
import tkinter as tk
from tkinter import ttk
import sqlite3


class School():
    '''
    School instance with name, abbr and goods.
    '''

    def __init__(self, name: str, abbr: str, route: str):
        self.name = name
        self.abbr = abbr
        self.route = route
        self.goods = {}  # (name, spec): number

    def add_goods(self, name: str, spec: str, number: float):
        k = (name, spec)
        self.goods[k] = self.goods.setdefault(k, 0) + number


class Route():
    def __init__(self, name: str):
        self.name = name
        self.schools = set()  # School


class RouteMap():
    '''
    Maintains a set of `School`s, while a set of `Route`s
    pointing to corresponding `School`s.
    '''

    def __init__(self):
        self.routes = {}   # str: Route
        self.schools = {}  # str: School

    def add_route(self, route: str, school: str, abbr: str):
        if route not in self.routes:
            self.routes[route] = Route(route)
        if school in self.schools:
            raise KeyError('线路分配表中不得出现重复的学校“{}”'.format(school))
        new_school = School(school, abbr, route)
        self.schools[school] = new_school
        self.routes[route].schools.add(new_school)


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


if __name__ == '__main__':
    pass
