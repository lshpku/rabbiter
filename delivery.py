#
# Delivery destination distributor.
#
import xlrd
from xlrd.sheet import Sheet
import xlwt
from xlwt import Worksheet, easyxf
import utils
from utils import RouteMap, AccountDatabase
from typing import Iterable

ALL_BORDER = 'border: left thin, right thin, top thin, bottom thin;'
FONT_240 = 'font: height 240;'

TEXT = easyxf(ALL_BORDER + FONT_240)
SCHOOL = easyxf(ALL_BORDER + FONT_240 +
                'align: wrap on, vert center, horz center;')

HEAD = easyxf('align: vert center, horz center;')
KIND = easyxf('pattern: pattern solid, fore_color light_green;')


def check_date_get_path(db: AccountDatabase) -> str:
    cur = db.select('DISTINCT DATE')
    dates = [i[0] for i in cur]
    if len(dates) == 0:
        raise KeyError('未发现数据或不能识别，请打开Excel文件并保存后重试')
    if len(dates) > 1:
        utils.log('建议一次只处理一天的数据，发现{}天'.format(len(dates)))
    date = xlrd.xldate_as_tuple(dates[0], 0)
    path = '分拣结果 {:02d}月{:02d}日.xls'.format(date[1], date[2])
    return path


def add_route(db: AccountDatabase, route_map: RouteMap):
    # add route to database
    db.add_colume('ROUTE CHAR(32)', 'ABBR CHAR(32)')
    stm = 'UPDATE TEMP SET ROUTE=?, ABBR=? WHERE SCHOOL=?'
    for school, (route, abbr) in route_map.schools.items():
        db.cur.execute(stm, (route, abbr, school))

    # ckeck for non-routed schools
    cur = db.select('DISTINCT SCHOOL, ROUTE')
    to_del = []
    for i in cur:
        if not i[1]:
            utils.log('学校“{}”无对应路线，已丢弃其所有记录'.format(i[0]))
            to_del.append(i[0])
    for i in to_del:
        db.cur.execute('DELETE FROM TEMP WHERE SCHOOL={}'.format(repr(i)))
    db.conn.commit()

    # add order
    routes = [None] * len(route_map.rt_idx)
    for r, i in route_map.rt_idx.items():
        routes[i] = r
    db.add_order('ROUTE', routes)
    schools = [None] * len(route_map.sc_idx)
    for s, i in route_map.sc_idx.items():
        schools[i] = s
    db.add_order('SCHOOL', schools)


def distribute(sheet: Worksheet, db: AccountDatabase, route_map: RouteMap):
    cur = db.select('DISTINCT ROUTE')
    route = cur.__next__()[0]
    sheet.write(0, 0, route, SCHOOL)

    # write goods tags in kinds' order
    kinds = db.sorted_one('KIND')
    goods = []
    for kind in kinds:
        cur = db.select('DISTINCT NAME', 'KIND={}'.format(repr(kind)))
        goods += [i[0] for i in cur]
    for i, gd in enumerate(goods):
        sheet.write(i+1, 0, gd, TEXT)
    goods_sum = [0] * len(goods)

    # write schools in appearing order
    cur = db.select('DISTINCT SCHOOL')
    schools = route_map.sort_school([i[0] for i in cur])
    for i, school in enumerate(schools):
        sheet.col(i+1).width_mismatch = True
        sheet.col(i+1).width = 2000
        sheet.write(0, i+1, route_map.schools[school][1], SCHOOL)
        for j, gd in enumerate(goods):
            where = 'SCHOOL={} AND NAME={}'.format(repr(school), repr(gd))
            cur = db.select('SUM(NUMBER)', where)
            value = cur.__next__()[0]
            goods_sum[j] += value if value else 0
            sheet.write(j+1, i+1, value if value else '', TEXT)

    # write goods sum
    sheet.col(i+2).width_mismatch = True
    sheet.col(i+2).width = 2000
    sheet.write(0, i+2, '总计', SCHOOL)
    for j, gd in enumerate(goods_sum):
        sheet.write(j+1, i+2, gd, TEXT)


def daily_sum(sheet: Worksheet, db: AccountDatabase):
    header = ['货品类别', '货品名称', '规格', '数量']
    for i, j in enumerate(header):
        sheet.write(0, i, j, HEAD)

    kinds = db.sorted_one('KIND')
    idx = 1
    for kind in kinds:
        sheet.write(idx, 0, kind, KIND)
        cur = db.select('DISTINCT NAME, SPEC', 'KIND={}'.format(repr(kind)))
        ns = [i for i in cur]
        for name, spec in ns:
            where = 'NAME={} AND SPEC={}'.format(repr(name), repr(spec))
            num = db.select('SUM(NUMBER)', where).__next__()[0]
            sheet.write(idx, 1, name)
            sheet.write(idx, 2, spec)
            sheet.write(idx, 3, num)
            idx += 1

    sheet.set_panes_frozen(True)
    sheet.set_horz_split_pos(1)


def plain_text(db: AccountDatabase) -> str:
    header = '线路\t客户名称\t货品名称\t数量\t规格\t单位'
    data = [header]
    for r in db.sorted_one('ROUTE'):
        for s in db.sorted_one('ABBR', 'ROUTE={}'.format(repr(r))):
            line = [r, s]
            tar = 'NAME, NUMBER, SPEC, UNIT'
            whr = 'ABBR={}'.format(repr(s))
            for name, num, spec, unit in db.select(tar, whr):
                nums = str(num)
                nums = nums[:-2] if nums.endswith('.0') else nums
                data.append('\t'.join(line + [name, nums, spec, unit]))
    return '\n'.join(data)


def handle(manifest: Sheet, route: Sheet, does: Iterable[str] = ...):
    db = AccountDatabase(manifest)
    path = check_date_get_path(db)
    workbook = xlwt.Workbook(encoding='utf-8')

    if does is ...:
        does = ['routing', 'daily', 'printer']
    does = set(does)

    if 'routing' in does:
        route_map = RouteMap(route)
        add_route(db, route_map)
        cur = db.select('DISTINCT ROUTE')
        routes = route_map.sort_route([i[0] for i in cur if i[0]])
        for route in routes:
            db.set_where('ROUTE={}'.format(repr(route)))
            distribute(workbook.add_sheet(route), db, route_map)
        db.set_where()

    if 'daily' in does:
        daily_sum(workbook.add_sheet('今日汇总'), db)

    if 'routing' in does or 'daily' in does:
        workbook.save(path)

    if 'printer' in does:
        text = plain_text(db)
        with open('标签打印表.txt', 'w', encoding='utf-8') as f:
            f.write(text)


if __name__ == '__main__':
    manifest = xlrd.open_workbook('data/众浩12月2日.xls')
    route = xlrd.open_workbook('data/线路分配表.xlsx')

    mani = manifest.sheet_by_index(0)
    rout = route.sheet_by_index(0)

    handle(mani, rout)
