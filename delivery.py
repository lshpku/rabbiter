#
# Delivery destination distributor.
#
import xlrd
from xlrd.sheet import Sheet
import xlwt
from xlwt import Worksheet
from utils import RouteMap, AccountDatabase


def text_style():
    border = xlwt.Borders()
    border.left = xlwt.Borders.THIN
    border.right = xlwt.Borders.THIN
    border.top = xlwt.Borders.THIN
    border.bottom = xlwt.Borders.THIN

    font = xlwt.Font()
    font.height = 240  # 12

    style = xlwt.XFStyle()
    style.borders = border
    style.font = font
    return style


def school_tag_style():
    border = xlwt.Borders()
    border.left = xlwt.Borders.THIN
    border.right = xlwt.Borders.THIN
    border.top = xlwt.Borders.THIN
    border.bottom = xlwt.Borders.THIN

    font = xlwt.Font()
    font.height = 240

    align = xlwt.Alignment()
    align.wrap = xlwt.Alignment.WRAP_AT_RIGHT
    align.vert = xlwt.Alignment.VERT_CENTER
    align.horz = xlwt.Alignment.HORZ_CENTER

    style = xlwt.XFStyle()
    style.borders = border
    style.font = font
    style.alignment = align
    return style


TEXT = text_style()
SCHOOL = school_tag_style()


def distribute(db: AccountDatabase, route_map: RouteMap):
    # get and check date
    cur = db.select('DISTINCT DATE')
    dates = [i[0] for i in cur]
    assert len(dates) == 1, '一次只能处理一天的数据'
    date = xlrd.xldate_as_tuple(dates[0], 0)
    path = '线路分拣表 {:02d}月{:02d}日.xls'.format(date[1], date[2])

    # add route to database
    db.cur.execute('ALTER TABLE TEMP ADD COLUMN ROUTE CHAR(32);')
    stm = 'UPDATE TEMP SET ROUTE={} WHERE SCHOOL={}'
    for school, (route, _) in route_map.schools.items():
        db.cur.execute(stm.format(repr(route), repr(school)))
    db.conn.commit()

    # ckeck for non-routed schools
    cur = db.select('DISTINCT SCHOOL, ROUTE')
    to_del = []
    for i in cur:
        if not i[1]:
            print('学校“{}”无对应路线，已丢弃其所有记录'.format(i[0]))
            to_del.append(i[0])
    for i in to_del:
        db.cur.execute('DELETE FROM TEMP WHERE SCHOOL={}'.format(repr(i)))

    workbook = xlwt.Workbook(encoding='utf-8')

    # draw each route
    cur = db.select('DISTINCT ROUTE')
    routes = route_map.sort_route([i[0] for i in cur if i[0]])
    for route in routes:
        sheet = workbook.add_sheet(route)
        sheet.write(0, 0, route, SCHOOL)

        # write goods tags
        stm = 'SELECT DISTINCT NAME FROM TEMP WHERE ROUTE={} AND KIND={}'
        goods = []
        for kind in ['肉', '菜', '油料干货', '调料制品']:
            cur = db.cur.execute(stm.format(repr(route), repr(kind)))
            goods += [i[0] for i in cur]
        for i, gd in enumerate(goods):
            sheet.write(i+1, 0, gd, TEXT)
        goods_sum = [0] * len(goods)

        # write schools
        cur = db.select('DISTINCT SCHOOL', 'ROUTE={}'.format(repr(route)))
        schools = route_map.sort_school([i[0] for i in cur])
        stm = 'SELECT SUM(NUMBER) FROM TEMP WHERE SCHOOL={} AND NAME={}'
        for i, school in enumerate(schools):
            sheet.col(i+1).width_mismatch = True
            sheet.col(i+1).width = 2000
            sheet.write(0, i+1, route_map.schools[school][1], SCHOOL)
            for j, gd in enumerate(goods):
                cur = db.cur.execute(stm.format(repr(school), repr(gd)))
                value = cur.__next__()[0]
                goods_sum[j] += value if value else 0
                sheet.write(j+1, i+1, value if value else '', TEXT)

        # write goods sum
        sheet.col(i+2).width_mismatch = True
        sheet.col(i+2).width = 2000
        sheet.write(0, i+2, '总计', SCHOOL)
        for j, gd in enumerate(goods_sum):
            sheet.write(j+1, i+2, gd, TEXT)

    workbook.save(path)


def handle(manifest: Sheet, route: Sheet):
    db = AccountDatabase(manifest)
    route_map = RouteMap(route)
    distribute(db, route_map)


if __name__ == '__main__':
    manifest = xlrd.open_workbook('data/众浩12月2日.xls')
    route = xlrd.open_workbook('data/线路分配表.xlsx')

    mani = manifest.sheet_by_index(0)
    rout = route.sheet_by_index(0)

    handle(mani, rout)
