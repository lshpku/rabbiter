import xlrd
from xlrd.sheet import Sheet
import xlwt
from xlwt import Worksheet
import os
import utils
from utils import AccountDatabase
import time

ARGS = [
    'DATE   CHAR(16) 1',   # 日期
    'CLIENT CHAR(32) 4',   # 单位全名
    'KIND   CHAR(16) 9',   # 产地
    'WEIGHT REAL     10',  # 销售数量
]

ARGS_CLI = [
    'SALES  CHAR(32) ',
    'CLIENT CHAR(32) ',
]


def monthly(sheet: Worksheet, db: AccountDatabase):
    y = db.distinct('YEAR').__next__()[0]
    m = db.distinct('MONTH').__next__()[0]
    d = max([i[0] for i in db.distinct('DAY')])
    sheet.write(0, 0, '{}年每月份客户销量汇总表'.format(y))
    sheet.write(0, 3, '{}-{:02d}-{:02d}'.format(y, m, d))
    pass


def salesman(sheet: Worksheet, db: AccountDatabase):
    s = db.distinct('SALES').__next__()[0]
    pass


def annually(sheet: Worksheet, db: AccountDatabase):
    pass


def handle(sheet: Sheet, log=print):
    db = AccountDatabase(sheet, ARGS, 10)
    workbook = xlwt.Workbook(encoding='utf-8')

    db.add_order('SALES', '唐曾谢蒙廖易')
    db.add_order('KIND', ['浓缩料', '乳猪料', '全价料', '禽料', '教槽'])

    # add y, m, d to table
    dates = [i[0] for i in db.distinct('DATE')]
    db.add_colume('YEAR INT', 'MONTH INT', 'DAY INT')
    for i in dates:
        t = time.strptime(i, '%Y-%m-%d')
        s = 'YEAR={}, MONTH={}, DAY={}'.format(t.tm_year, t.tm_mon, t.tm_mday)
        db.update(s, 'DATE={}'.format(repr(i)))
    db.conn.commit()

    # write monthly
    years = db.sorted_one('YEAR')
    for y in years:
        months = db.sorted_one('MONTH', 'YEAR={}'.format(y))
        for m in months:
            db.set_where('YEAR={} AND MONTH={}'.format(y, m))
            if len(years) > 1:  # multiple years
                label = '{:02d}年{}月'.format(y % 100, m)
            else:
                label = '{}月'.format(m)
            monthly(workbook.add_sheet(label), db)

    # write salesmen
    sales = db.sorted_one('SALES')
    for s in sales:
        db.set_where('SALES={}'.format(repr(s)))
        salesman(workbook.add_sheet(s), db)

    # write anually
    for y in years:
        db.set_where('YEAR={}'.format(y))
        annually(workbook.add_sheet('{}年总合计'.format(y)), db)

    cur = db.distinct('DAY', 'YEAR={} AND MONTH={}'.format(y, m))
    d = max([i[0] for i in cur])
    path = '{:02d}.{:02d}.{:02d}-销量邮件表.xls'.format(y % 100, m, d)
    workbook.save(path)


if __name__ == '__main__':
    workbook = xlrd.open_workbook(
        'data/单位销售明细数据-2020_02_05-11_50_19.xls')
    handle(workbook.sheet_by_index(0))
