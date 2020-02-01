#
# Delivery destination distributor.
#
import xlrd
from xlrd.sheet import Sheet
import xlwt
from xlwt import Worksheet
from utils import School, Route, RouteMap


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


def parse_routes(route_table: Sheet) -> RouteMap:
    '''
    Build route map with schools.
    '''
    routes = route_table.col_values(0)
    schools = route_table.col_values(1)
    assert len(schools) == len(routes)
    abbrs = route_table.col_values(2)
    assert len(abbrs) == len(schools)

    route_map = RouteMap()
    for i, route in enumerate(routes):
        try:
            route_map.add_route(route, schools[i], abbrs[i])
        except KeyError as e:
            print(e.args[0])
    return route_map


def parse_manifest(manifest: Sheet, route_map: RouteMap):
    '''
    Map items in manifest to schools in route map.
    '''
    schools = manifest.col_values(2, 1)
    names = manifest.col_values(6, 1)
    assert len(names) == len(schools)
    specs = manifest.col_values(7, 1)
    assert len(specs) == len(names)
    numbers = manifest.col_values(8, 1)
    assert len(numbers) == len(specs)

    for i, school in enumerate(schools):
        if school not in route_map.schools:
            print('学校“{}”未出现在线路分配表中'.format(school))
            continue
        route_map.schools[school].add_goods(names[i], specs[i], numbers[i])


def save_result(route_map: RouteMap, path: str):
    workbook = xlwt.Workbook(encoding='utf-8')

    for name, route in route_map.routes.items():
        sheet = workbook.add_sheet(name)

        goods_sum = {}  # (name, spec): number
        for school in route.schools:
            for k, w in school.goods.items():
                if k in goods_sum:
                    goods_sum[k] += w
                else:
                    goods_sum[k] = w

        sheet.write(0, 0, route.name, SCHOOL)

        # write goods tags
        goods_idx = {}  # (name, spec): index
        for i, goods in enumerate(goods_sum.keys()):
            goods_idx[goods] = i+1
            sheet.write(i+1, 0, goods[0], TEXT)

        # write schools
        for i, school in enumerate(route.schools):
            sheet.col(i+1).width_mismatch = True
            sheet.col(i+1).width = 2000
            sheet.write(0, i+1, school.abbr, SCHOOL)
            for goods, index in goods_idx.items():
                if goods in school.goods:
                    sheet.write(index, i+1, school.goods[goods], TEXT)
                else:
                    sheet.write(index, i+1, '', TEXT)

        # write goods sum
        sheet.col(i+2).width_mismatch = True
        sheet.col(i+2).width = 2000
        sheet.write(0, i+2, '总计', SCHOOL)
        for k, w in goods_sum.items():
            sheet.write(goods_idx[k], len(route.schools)+1, w, TEXT)

    workbook.save(path)


def distribute(manifest: Sheet, route: Sheet, save_path: str):
    route_map = parse_routes(route)
    parse_manifest(manifest, route_map)
    save_result(route_map, save_path)


if __name__ == '__main__':
    manifest = xlrd.open_workbook('data/12月 008仓库 开票明细.xls')
    route = xlrd.open_workbook('data/线路分配表.xlsx')

    mani = manifest.sheet_by_index(0)
    rout = route.sheet_by_index(0)

    distribute(mani, rout, 'result.xls')
