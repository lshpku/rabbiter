#
# PivotTable Generator.
#
# P2d:
#   日期 - 餐类
# P2d2:
#   货品类别 - 餐类
#   货品名称 - 餐类
# P3d:
#   货品名称 - 货品类别 - 日期
#
import xlrd
from xlrd.sheet import Sheet
import xlwt
from xlwt import Worksheet, easyxf
from utils import AccountDatabase, RouteMap
import os
import sqlite3

BORDER_FULL = 'border: left thin, right thin, top thin, bottom thin;'
HORZ_CENTER = 'align: horz center;'
BKG_GREEN = 'pattern: pattern solid, fore_color light_green;'

DATE = easyxf(BORDER_FULL, 'yyyy/m/d')
TEXT = easyxf(BORDER_FULL + 'align: wrap on;')
HEAD = easyxf(BORDER_FULL + HORZ_CENTER)
MONTH = easyxf(BORDER_FULL + HORZ_CENTER, 'm/d')
TITLE = easyxf(HORZ_CENTER)
SUM = easyxf(BORDER_FULL + BKG_GREEN)


def pivottable_2d(sheet: Worksheet, db: AccountDatabase,
                  vert: str, horz: str):
    ys = db.sorted_one(vert)
    xs = db.sorted_one(horz)

    for i, x in enumerate(xs):
        sheet.write(0, i+1, x, HEAD)
    sheet.write(0, i+2, '总计', HEAD)
    sheet.write(0, 0, '', HEAD)

    col_sum = [0] * len(xs)
    for i, y in enumerate(ys):
        sheet.write(i+1, 0, y, DATE)
        row_sum = 0
        for j, x in enumerate(xs):
            where = '{}={} AND {}={}'.format(vert, repr(y), horz, repr(x))
            bsum = db.select('SUM(TOTAL)', where).__next__()[0]
            sheet.write(i+1, j+1, bsum if bsum else '', TEXT)
            row_sum += bsum if bsum else 0
            col_sum[j] += bsum if bsum else 0
        sheet.write(i+1, j+2, row_sum, TEXT)

    sheet.write(i+2, 0, '总计', TEXT)
    for j, s in enumerate(col_sum):
        sheet.write(i+2, j+1, s, TEXT)
    sheet.write(i+2, j+2, sum(col_sum), TEXT)


def pivottable_2d2(sheet: Worksheet, db: AccountDatabase,
                   vert: str, horz: str):
    ys = db.sorted_one(vert)
    xs = db.sorted_one(horz)

    for i, x in enumerate(xs):
        sheet.write_merge(0, 0, i*2+1, i*2+2, x, HEAD)
        sheet.write(1, i*2+1, '数量', HEAD)
        sheet.write(1, i*2+2, '货款', HEAD)
    sheet.write_merge(0, 1, 0, 0, '', HEAD)

    col_sum1 = [0] * len(xs)
    col_sum2 = [0] * len(xs)

    for i, y in enumerate(ys):
        sheet.write(i+2, 0, y, TEXT)
        for j, x in enumerate(xs):
            where = '{}={} AND {}={}'.format(vert, repr(y), horz, repr(x))
            sum1 = db.select('SUM(NUMBER)', where).__next__()[0]
            sheet.write(i+2, j*2+1, sum1 if sum1 else '', TEXT)
            col_sum1[j] += sum1 if sum1 else 0

            sum2 = db.select('SUM(TOTAL)', where).__next__()[0]
            sheet.write(i+2, j*2+2, sum2 if sum2 else '', TEXT)
            col_sum2[j] += sum2 if sum2 else 0

    sheet.write(i+3, 0, '总计', TEXT)
    for j, s in enumerate(col_sum1):
        sheet.write(i+3, j*2+1, s, TEXT)
    for j, s in enumerate(col_sum2):
        sheet.write(i+3, j*2+2, s, TEXT)


def pivottable_3d(sheet: Worksheet, db: AccountDatabase):
    dates = db.sorted_one('DATE')
    month = xlrd.xldate_as_tuple(dates[0], 0)[1]

    meal = db.select('DISTINCT MEAL').__next__()[0]
    school = db.select('DISTINCT SCHOOL').__next__()[0]
    title = '横县农村义务教育学生营养改善计划每日开餐情况统计 （{}月{}）'
    title = title.format(month, meal)
    sheet.write_merge(0, 0, 0, len(dates)+1, title, TITLE)
    sheet.write(1, 0, '学校名称：{}'.format(school))
    sheet.write(2, 0, '明细', HEAD)

    for i, date in enumerate(dates):
        xld = xlrd.xldate_as_tuple(date, 0)[1]
        sheet.write(2, i+1, date, MONTH)
        sheet.col(i+1).width_mismatch = True
        sheet.col(i+1).width = 1500
    sheet.write(2, i+2, '总计', HEAD)

    where = 'DATE={} AND NAME={}'
    idx = 3

    sheet.write(idx, 0, '大米', SUM)
    row_sum = 0
    for i, date in enumerate(dates):
        cur = db.select('SUM(TOTAL)', where.format(date, repr('大米')))
        xs = cur.__next__()[0]
        row_sum += xs if xs else 0
        sheet.write(idx, i+1, xs if xs else '', SUM)
    sheet.write(idx, i+2, row_sum, SUM)
    idx += 1

    kinds = db.sorted_one('KIND')
    for kind in kinds:
        names = db.sorted_one('NAME', 'KIND={}'.format(repr(kind)))
        col_sum = [0] * len(dates)
        for name in names:
            row_sum = 0
            sheet.write(idx, 0, name, TEXT)
            for i, date in enumerate(dates):
                cur = db.select('SUM(TOTAL)', where.format(date, repr(name)))
                xs = cur.__next__()[0]
                row_sum += xs if xs else 0
                col_sum[i] += xs if xs else 0
                sheet.write(idx, i+1, xs if xs else '', TEXT)
            sheet.write(idx, i+2, row_sum, TEXT)
            idx += 1
        sheet.write(idx, 0, '{}合计'.format(kind), SUM)
        for i, s in enumerate(col_sum):
            sheet.write(idx, i+1, s, SUM)
        sheet.write(idx, i+2, sum(col_sum), SUM)
        idx += 1

    for i in ['每日合计', '每日开餐人数', '人均开餐金额', '陪餐人数']:
        sheet.write(idx, 0, i, TEXT)
        for j in range(1, len(dates)+2):
            sheet.write(idx, j, '', TEXT)
        idx += 1


def handle_school(school: str, db: AccountDatabase, save_path: str):
    '''
    Save one school's data in one workbook.
    '''
    db.set_where('SCHOOL={}'.format(repr(school)))
    workbook = xlwt.Workbook(encoding='utf-8')

    pivottable_2d(workbook.add_sheet('日期'), db, 'DATE', 'MEAL')

    pivottable_2d2(workbook.add_sheet('类别'), db, 'KIND', 'MEAL')
    pivottable_2d2(workbook.add_sheet('品种'), db, 'NAME', 'MEAL')

    meals = db.sorted_one('MEAL')
    for meal in meals:
        stm = 'SCHOOL={} AND MEAL={}'
        db.set_where(stm.format(repr(school), repr(meal)))
        pivottable_3d(workbook.add_sheet('{}每日用餐'.format(meal)), db)

    workbook.save(save_path)


def handle(sheet: Sheet, log=print):
    '''
    Handle accounts of all schools in a month.
    '''
    log('读入表单数据')
    db = AccountDatabase(sheet)
    cur = db.select('DISTINCT SCHOOL')
    schools = [i[0] for i in cur if i[0]]
    cur = db.select('DISTINCT DATE')
    dates = [i[0] for i in cur if i[0]]
    months = set([xlrd.xldate_as_tuple(i, 0)[1] for i in dates])
    assert len(months) == 1, '一次只能处理一个月的数据'

    month = '{:02d}月'.format([i for i in months][0])
    path = os.path.join('.', month)
    if not os.path.exists(path):
        os.mkdir(path)
    for i, school in enumerate(schools):
        fpath = os.path.join(path, '{} {}.xls'.format(school, month))
        log('{}/{}：{}'.format(i+1, len(schools), school))
        handle_school(school, db, fpath)


if __name__ == '__main__':
    workbook = xlrd.open_workbook('data/12月 008仓库 开票明细.xls')
    handle(workbook.sheet_by_index(0))
