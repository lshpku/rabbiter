import xlrd
from xlrd.sheet import Sheet
import xlwt
from xlwt import Worksheet, easyxf
import os
import utils
from utils import AccountDatabase, log
import time
from typing import Dict

ARGS = [
    'DATE   CHAR(16) 1',  # 日期
    'CLIENT CHAR(32) 3',  # 单位全名
    'KIND   CHAR(16) 5',  # 产地
    'WEIGHT REAL     8',  # 销售数量
]

KINDS = ['浓缩料', '乳猪料', '全价料', '禽料', '教槽', '预混料']
EXPANDS = ['9018', '611']
SALESMEN = '唐曾谢蒙廖易'
HEADER = ['单位全名', '地区'] + KINDS + ['总计'] + EXPANDS

BORDER_FULL = 'border: left thin, right thin, top thin, bottom thin;'
HORZ_CENTER = 'align: horz center;'
BKG_GREEN = 'pattern: pattern solid, fore_color light_green;'

BORD = easyxf(BORDER_FULL)
BORDC = easyxf(BORDER_FULL + HORZ_CENTER)
UBORDC = easyxf(HORZ_CENTER)


def write_sales(sheet: Worksheet, db: AccountDatabase,
                sales: str, nrow: int) -> int:
    ssum = [0] * (len(HEADER)-2)
    for cli in db.sales_map[sales]:
        sheet.write(nrow, 0, cli, BORD)
        sheet.write(nrow, 1, sales, BORDC)

        # write basic kinds
        ksum = 0
        for i, kind in enumerate(KINDS, 2):
            whr = 'CLIENT={} AND KIND={}'.format(repr(cli), repr(kind))
            val = db.select('SUM(WEIGHT)', whr).__next__()[0]
            val = val/1000 if val else 0
            ksum += val
            ssum[i-2] += val
            sheet.write(nrow, i, val if val else '', BORD)
        sheet.write(nrow, i+1, ksum if ksum else '', BORD)
        ssum[i-1] += ksum

        # write extra kinds
        for i, kind in enumerate(EXPANDS, i+2):
            try:
                kind = '{}.0'.format(int(kind))
            except:
                pass
            whr = 'CLIENT={} AND KIND={}'.format(repr(cli), repr(kind))
            val = db.select('SUM(WEIGHT)', whr).__next__()[0]
            val = val/1000 if val else 0
            ssum[i-2] += val
            sheet.write(nrow, i, val if val else '', BORD)
        nrow += 1

    # write sales sum
    sheet.write(nrow, 1, '合计', UBORDC)
    for i, val in enumerate(ssum, 2):
        sheet.write(nrow, i, val)
    nrow += 2
    return nrow


def monthly(sheet: Worksheet, db: AccountDatabase):
    y = db.distinct('YEAR').__next__()[0]
    header = (['单位全名', '邮编'] +
              ['{}月合计'.format(i+1) for i in range(12)] +
              ['年合计'])
    for i, j in enumerate(header):
        sheet.write(0, i, j, BORDC)
        sheet.col(i).width_mismatch = True
        sheet.col(i).width = 3400 if i < 1 else 2100

    clients = []
    for cli in db.sales_map.values():
        clients += cli
    clients.sort()

    for i, cli in enumerate(clients, 1):
        sheet.write(i, 0, cli, BORD)
        sheet.write(i, 1, db.client_map[cli], BORDC)
        rsum = 0
        for j in range(1, 13):
            whr = 'MONTH={} AND CLIENT={} AND BASIC=1'.format(j, repr(cli))
            val = db.select('SUM(WEIGHT)', whr).__next__()[0]
            val = val/1000 if val else 0
            sheet.write(i, j+1, val if val else '', BORD)
            rsum += val if val else 0
        sheet.write(i, 14, rsum, BORD)


def salesman(sheet: Worksheet, db: AccountDatabase):
    s = db.distinct('SALES').__next__()[0]
    clients = db.sorted_one('CLIENT')

    for i, j in enumerate(HEADER):
        sheet.write(0, i, j, BORDC)
        sheet.col(i).width_mismatch = True
        sheet.col(i).width = 3400 if i < 1 else 2100
    nrow = 1

    write_sales(sheet, db, s, nrow)


def annually(sheet: Worksheet, db: AccountDatabase):
    sheet.write(0, 0, '', BORD)
    header = HEADER[2:]
    for i, h in enumerate(header, 1):
        sheet.write(0, i, h, BORDC)
    for i in range(len(header)+1):
        sheet.col(i).width_mismatch = True
        sheet.col(i).width = 2100

    ssum = [0] * len(header)
    for month in range(1, 13):
        sheet.write(month, 0, '{}月'.format(month), BORD)
        # write basic kinds
        ksum = 0
        for i, kind in enumerate(KINDS, 1):
            whr = 'MONTH={} AND KIND={}'.format(month, repr(kind))
            val = db.select('SUM(WEIGHT)', whr).__next__()[0]
            val = val/1000 if val else 0
            ksum += val
            ssum[i-1] += val
            sheet.write(month, i, val if val else '', BORD)
        sheet.write(month, i+1, ksum if ksum else '', BORD)
        ssum[i] += ksum

        # write extra kinds
        for i, kind in enumerate(EXPANDS, i+2):
            try:
                kind = '{}.0'.format(int(kind))
            except:
                pass
            whr = 'MONTH={} AND KIND={}'.format(month, repr(kind))
            val = db.select('SUM(WEIGHT)', whr).__next__()[0]
            val = val/1000 if val else 0
            ssum[i-1] += val
            sheet.write(month, i, val if val else '', BORD)

    # write col sum
    sheet.write(13, 0, '合计', UBORDC)
    for i, val in enumerate(ssum, 1):
        sheet.write(13, i, val)


def make_client_map(client_list: Sheet) -> Dict[str, str]:
    clients = client_list.col_values(1, 1)
    sales = client_list.col_values(3, 1)
    client_map = {}
    for i, client in enumerate(clients):
        sal = sales[i]
        ret = client_map.setdefault(client, sal)
        if ret != sal:
            log('“{}”同时属于“{}”和“{}”，自动归为“{}”'
                .format(client, ret, sal, ret))
    return client_map


def handle(manifest: Sheet, client_list: Sheet, does_sales: bool,
           does_annually: bool, does_monthly: bool):
    db = AccountDatabase(manifest, ARGS, 10)
    workbook = xlwt.Workbook(encoding='utf-8')

    db.add_order('SALES', SALESMEN)
    db.add_order('KIND', KINDS)

    # add y, m, d to table
    dates = [i[0] for i in db.distinct('DATE')]
    db.add_colume('YEAR INT', 'MONTH INT', 'DAY INT')
    for i in dates:
        t = time.strptime(i, '%Y-%m-%d')
        s = 'YEAR={}, MONTH={}, DAY={}'.format(t.tm_year, t.tm_mon, t.tm_mday)
        db.update(s, 'DATE={}'.format(repr(i)))

    years = db.sorted_one('YEAR')
    months = db.sorted_one('MONTH', 'YEAR={}'.format(years[-1]))

    # add basic_flag to table
    db.add_colume('BASIC INT')
    for i in KINDS:
        db.update('BASIC=1', 'KIND={}'.format(repr(i)))

    # add sales_map to table
    if does_sales or does_monthly:
        db.add_colume('SALES CHAR(16)')
        db.client_map = make_client_map(client_list)
        clients = [i[0] for i in db.distinct('CLIENT')]
        for cli in clients:
            sal = db.client_map.get(cli)
            if sal is None:
                log('客户“{}”无对应业务员，已丢弃'.format(cli))
            db.update('SALES={}'.format(repr(sal)),
                      'CLIENT={}'.format(repr(cli)))

        db.sales_map = {}
        for cli, sal in db.client_map.items():
            db.sales_map.setdefault(sal, []).append(cli)
        for clis in db.sales_map.values():
            clis.sort()

    # write salesmen
    if does_sales:
        sales = db.sorted_one('SALES')
        for s in sales:
            db.set_where('SALES={}'.format(repr(s)))
            salesman(workbook.add_sheet(s if s else '其他'), db)

    # write anually
    if does_annually:
        for y in years:
            db.set_where('YEAR={}'.format(y))
            annually(workbook.add_sheet('{}总(料型)'.format(y)), db)

    # write monthly
    if does_monthly:
        for y in years:
            db.set_where('YEAR={}'.format(y))
            monthly(workbook.add_sheet('{}总(客户)'.format(y)), db)

    y, m = years[-1], months[-1]
    cur = db.distinct('DAY', 'YEAR={} AND MONTH={}'.format(y, m))
    d = max([i[0] for i in cur])
    path = '{:02d}.{:02d}.{:02d}-销量邮件表.xls'.format(y % 100, m, d)
    workbook.save(path)


if __name__ == '__main__':
    workbook = xlrd.open_workbook(
        'data/单位销售明细数据-2020_02_13-13_14_53.xls')
    client_list = xlrd.open_workbook('data/客户名.xls')
    handle(workbook.sheet_by_index(0), client_list.sheet_by_index(0),
           True, True, True)
