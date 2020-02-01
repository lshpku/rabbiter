#
# PivotTable Generator.
#
# 日期 - 餐类
# 货品类别 - 餐类
# 货品名称 - 餐类
# 货品名称 - 货品类别 - 日期
#
import xlrd
from xlrd.sheet import Sheet
import xlwt
from xlwt import Worksheet
from utils import School, Route, RouteMap, AccountDatabase
import os
import sqlite3


def full_border():
    border = xlwt.Borders()
    border.left = xlwt.Borders.THIN
    border.right = xlwt.Borders.THIN
    border.top = xlwt.Borders.THIN
    border.bottom = xlwt.Borders.THIN
    return border


def date_style():
    style = xlwt.XFStyle()
    style.borders = full_border()
    style.num_format_str = 'yyyy/m/d'
    return style


def text_style():
    align = xlwt.Alignment()
    align.wrap = xlwt.Alignment.WRAP_AT_RIGHT
    style = xlwt.XFStyle()
    style.borders = full_border()
    style.alignment = align
    return style


def head_style():
    align = xlwt.Alignment()
    align.horz = xlwt.Alignment.HORZ_CENTER
    style = xlwt.XFStyle()
    style.borders = full_border()
    style.alignment = align
    return style


def month_style():
    align = xlwt.Alignment()
    align.horz = xlwt.Alignment.HORZ_CENTER
    style = xlwt.XFStyle()
    style.borders = full_border()
    style.alignment = align
    style.num_format_str = 'm/d'
    return style


def title_style():
    align = xlwt.Alignment()
    align.horz = xlwt.Alignment.HORZ_CENTER
    style = xlwt.XFStyle()
    style.alignment = align
    return style


def sum_style():
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = xlwt.Style.colour_map['light_green']
    style = xlwt.XFStyle()
    style.pattern = pattern
    style.borders = full_border()
    return style


DATE = date_style()
TEXT = text_style()
HEAD = head_style()
MONTH = month_style()
TITLE = title_style()
SUM = sum_style()
SCHOOL = None


def pivottable_2d(sheet: Worksheet, db: AccountDatabase,
                  vert: str, horz: str):
    stm = 'SELECT DISTINCT {} FROM TEMP WHERE SCHOOL={}'
    cur = db.cur.execute(stm.format(vert, SCHOOL))
    ys = [i[0] for i in cur]
    cur = db.cur.execute(stm.format(horz, SCHOOL))
    xs = [i[0] for i in cur]

    for i, x in enumerate(xs):
        sheet.write(0, i+1, x, HEAD)
    sheet.write(0, i+2, '总计', HEAD)
    sheet.write(0, 0, '', HEAD)

    col_sum = [0] * len(xs)
    for i, y in enumerate(ys):
        sheet.write(i+1, 0, y, DATE)
        row_sum = 0
        for j, x in enumerate(xs):
            stm = ('SELECT SUM(TOTAL) FROM TEMP WHERE'
                   ' {}={} AND {}={} AND SCHOOL={}')
            stm = stm.format(vert, repr(y), horz, repr(x), SCHOOL)
            bsum = db.cur.execute(stm).__next__()[0]
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
    stm = 'SELECT DISTINCT {} FROM TEMP WHERE SCHOOL={}'
    cur = db.cur.execute(stm.format(vert, SCHOOL))
    ys = [i[0] for i in cur]
    cur = db.cur.execute(stm.format(horz, SCHOOL))
    xs = [i[0] for i in cur]

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
            stm = ('SELECT SUM({}) FROM TEMP WHERE'
                   ' {}={} AND {}={} AND SCHOOL={}')
            stm1 = stm.format('NUMBER', vert, repr(y), horz, repr(x), SCHOOL)
            sum1 = db.cur.execute(stm1).__next__()[0]
            sheet.write(i+2, j*2+1, sum1 if sum1 else '', TEXT)
            col_sum1[j] += sum1 if sum1 else 0

            stm2 = stm.format('TOTAL', vert, repr(y), horz, repr(x), SCHOOL)
            sum2 = db.cur.execute(stm2).__next__()[0]
            sheet.write(i+2, j*2+2, sum2 if sum2 else '', TEXT)
            col_sum2[j] += sum2 if sum2 else 0

    sheet.write(i+3, 0, '总计', TEXT)
    for j, s in enumerate(col_sum1):
        sheet.write(i+3, j*2+1, s, TEXT)
    for j, s in enumerate(col_sum2):
        sheet.write(i+3, j*2+2, s, TEXT)


def pivottable_3d(sheet: Worksheet, db: AccountDatabase):
    stm = 'SELECT DISTINCT DATE FROM TEMP WHERE SCHOOL={}'
    cur = db.cur.execute(stm.format(SCHOOL))
    dates = [i[0] for i in cur]
    dates.sort()

    month = xlrd.xldate_as_tuple(dates[0], 0)[1]

    title = '横县农村义务教育学生营养改善计划每日开餐情况统计 （{}月营养餐）'
    sheet.write_merge(0, 0, 0, len(dates)+1, title.format(month), TITLE)
    sheet.write(1, 0, '学校名称：{}'.format(SCHOOL.strip("'")))
    sheet.write(2, 0, '明细', HEAD)

    for i, date in enumerate(dates):
        xld = xlrd.xldate_as_tuple(date, 0)[1]
        sheet.write(2, i+1, date, MONTH)
        sheet.col(i+1).width_mismatch = True
        sheet.col(i+1).width = 1500
    sheet.write(2, i+2, '总计', HEAD)

    stm = 'SELECT SUM(TOTAL) FROM TEMP WHERE DATE={} AND NAME={} AND SCHOOL={}'
    idx = 3

    # 大米
    sheet.write(idx, 0, '大米', SUM)
    row_sum = 0
    for i, date in enumerate(dates):
        cur = db.cur.execute(stm.format(date, repr('大米'), SCHOOL))
        xs = cur.__next__()[0]
        row_sum += xs if xs else 0
        sheet.write(idx, i+1, xs if xs else '', SUM)
    sheet.write(idx, i+2, row_sum, SUM)
    idx += 1

    stm_kind = 'SELECT DISTINCT NAME FROM TEMP WHERE KIND={} AND SCHOOL={}'

    for kinds, tag in [(['菜'], '菜'), (['肉'], '肉'),
                       (['油料干货', '调料制品'], '调料')]:
        names = []
        for i in kinds:
            cur = db.cur.execute(stm_kind.format(repr(i), SCHOOL))
            names += [i[0] for i in cur]
        col_sum = [0] * len(dates)
        for name in names:
            row_sum = 0
            sheet.write(idx, 0, name, TEXT)
            for i, date in enumerate(dates):
                cur = db.cur.execute(stm.format(date, repr(name), SCHOOL))
                xs = cur.__next__()[0]
                row_sum += xs if xs else 0
                col_sum[i] += xs if xs else 0
                sheet.write(idx, i+1, xs if xs else '', TEXT)
            sheet.write(idx, i+2, row_sum, TEXT)
            idx += 1
        sheet.write(idx, 0, '{}合计'.format(tag), SUM)
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
    global SCHOOL
    SCHOOL = repr(school)
    workbook = xlwt.Workbook(encoding='utf-8')
    pivottable_2d(workbook.add_sheet('日期'), db, 'DATE', 'MEAL')
    pivottable_2d2(workbook.add_sheet('类别'), db, 'KIND', 'MEAL')
    pivottable_2d2(workbook.add_sheet('品种'), db, 'NAME', 'MEAL')
    pivottable_3d(workbook.add_sheet('营养餐每日用餐'), db)
    workbook.save(save_path)


def handle_sheet(sheet: Sheet, save_path='.', log=print):
    '''
    Handle accounts of all schools in a month.
    '''
    log('读入表单数据')
    db = AccountDatabase(sheet)
    cur = db.cur.execute('SELECT DISTINCT SCHOOL FROM TEMP')
    schools = [i[0] for i in cur if i[0]]
    cur = db.cur.execute('SELECT DISTINCT DATE FROM TEMP')
    dates = [i[0] for i in cur if i[0]]
    months = set([xlrd.xldate_as_tuple(i, 0)[1] for i in dates])
    assert len(months) == 1, '一次只能处理一个月的数据'

    month = '{:02d}月'.format([i for i in months][0])
    path = os.path.join(save_path, month)
    if not os.path.exists(path):
        os.mkdir(path)
    for i, school in enumerate(schools):
        fpath = os.path.join(path, '{} {}.xls'.format(school, month))
        log('{}/{}：{}'.format(i+1, len(schools), school))
        handle_school(school, db, fpath)


if __name__ == '__main__':
    workbook = xlrd.open_workbook('data/12月 008仓库 开票明细(all).xls')
    handle_sheet(workbook.sheet_by_index(0), '.')
