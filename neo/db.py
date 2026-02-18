import xlrd
from xlrd.sheet import Sheet
import xlwt
from xlwt import Worksheet, easyxf
from datetime import datetime
from collections import defaultdict
from typing import Any, Callable, Dict, List, Tuple


class Table:
    def __init__(self, keys: Dict[str, int], values: List[List[Any]]):
        self.keys = keys
        self.values = values

    def col(self, key: str, unique=False):
        """获取key对应的整列"""
        col_idx = self.keys[key]
        values = [row[col_idx] for row in self.values]
        if not unique:
            return values
        return list(dict.fromkeys(values))

    def select_eq(self, key: str, value: Any):
        """获取key所在列的值为value的所有行构成的子表"""
        col_idx = self.keys[key]
        values = [row for row in self.values if row[col_idx] == value]
        return Table(self.keys, values)
    
    def select_fn(self, key: str, match: Callable[[Any], bool]):
        col_idx = self.keys[key]
        values = [row for row in self.values if match(row[col_idx])]
        return Table(self.keys, values)

    def __bool__(self):
        return bool(self.values)


class SalesDatabase(Table):
    required_keys = ["制单日期", "单位全名", "料型", "销售数量"]

    def __init__(self, sheet: Sheet):
        keys, values = {}, []

        # 读取表头
        for i in range(0, sheet.nrows):
            row = sheet.row_values(i)
            if not (row and row[0] == "行号"):
                continue
            for key in self.required_keys:
                assert key in row, f"“明细数据表”缺失列：{key}"
            for j, key in enumerate(row):
                old_j = keys.setdefault(key, j)
                assert old_j == j, f"禁止出现重复列（第{j}列和第{keys[key]}列）：{key}"
            break

        # 读取内容
        for i in range(i + 1, sheet.nrows):
            row = sheet.row_values(i)
            if (len(row) > 1) and (not row[1]):  # 最后一行
                continue
            values.append(row)

        # 转换类型
        date_idx = keys.get("制单日期")
        for row in values:
            y, m, d = row[date_idx].split("-")
            row[date_idx] = (int(y), int(m), int(d))

        self.keys = keys
        self.values = values


class RebateDatabase(Table):
    required_keys = ["单位全名", "湖大", "中转站销量", "费用"]

    def __init__(self, sheet: Sheet):
        header_row = sheet.row_values(0)
        for key in self.required_keys:
            assert key in header_row, f"“返利”表缺失列：{key}"

        self.keys = {key: i for i, key in enumerate(header_row) if key}
        self.values = [
            sheet.row_values(i)[:len(self.keys)]
            for i in range(1, sheet.nrows) 
        ]


class KindOrder:
    def __init__(self, sheet: Sheet):
        self.order_map = {}
        self.counted = {}
        for i in range(1, sheet.nrows):
            row = sheet.row_values(i)
            kind = row[0]
            counted = row[1] == "是"
            if kind:
                self.order_map[kind] = i
                self.counted[kind] = counted

    def sort(self, keys: List[str]):
        key_map = {
            i: k for k in keys
            if (i := self.order_map.get(k)) is not None
        }        
        return [item[1] for item in sorted(key_map.items())]


class PriceFactor:
    required_keys = ["品种", "折合量系数", "元"]

    def __init__(self, sheet: Sheet):
        header_row = sheet.row_values(0)
        for key in self.required_keys:
            assert key in header_row, f"“费用系数”表缺失列：{key}"

        key_map = {key: i for i, key in enumerate(header_row) if key}
        kind_i = key_map["品种"]
        factor_i = key_map["折合量系数"]
        price_i = key_map["元"]

        factors, prices = {}, {}
        for i in range(1, sheet.nrows):
            row = sheet.row_values(i)
            kind = row[kind_i]
            if kind:
                factor, price = row[factor_i], row[price_i]
                factors[kind] = factor if isinstance(factor, (int, float)) else 0
                prices[kind] = price if isinstance(price, (int, float)) else 0

        self.factors = factors
        self.prices = prices


def read_rebate_book(path) -> Tuple[RebateDatabase, KindOrder, PriceFactor]:
    required_sheets = {
        "返利": RebateDatabase,
        "料型排序": KindOrder,
        "费用系数": PriceFactor,
    }
    outputs = {key: None for key in required_sheets}

    book = xlrd.open_workbook(path)

    for sheet in book.sheets():
        if init_fn := required_sheets.get(sheet.name):
            print(f"打开表：{sheet.name}")
            outputs[sheet.name] = init_fn(sheet)

    missing_keys = [key for key, value in outputs.items() if value is None]
    assert not missing_keys, f"返利文档中缺少以下表格：{'、'.join(missing_keys)}"

    return tuple(outputs.values())


BORDER_FULL = 'border: left thin, right thin, top thin, bottom thin;'
HORZ_CENTER = 'align: horz center;'
VERT_TOP = 'align: vert top;'
TEXT_RED = "font: colour red;"

XF_BOX = easyxf(BORDER_FULL)
XF_BOX_C = easyxf(BORDER_FULL + HORZ_CENTER)
XF_BOX_TC = easyxf(BORDER_FULL + VERT_TOP + HORZ_CENTER)
XF_BOX_RED = easyxf(BORDER_FULL + TEXT_RED)
XF_BOX_C_RED = easyxf(BORDER_FULL + HORZ_CENTER + TEXT_RED)

QUANT = lambda x: x / 1000 if isinstance(x, (int, float)) else x

rebate_db, kind_order, price_factor = None, None, None
area_ref = None
now_str = None


def write_area_sheet(sheet: Worksheet, area: str, sales_db: SalesDatabase):
    print(f"绘制区域：{area}")

    # 写header并调整列宽
    header = ["单位简称", "料型", *(f"{i + 1}月" for i in range(12)), "总计"]
    for i, text in enumerate(header):
        sheet.write(0, i, text, XF_BOX_C)
        sheet.col(i).width_mismatch = True
        sheet.col(i).width = 2048 if i < 1 else 1530

    # 写每个客户
    row_idx = 1
    global_monthly_sales = defaultdict(int)
    for client_full_name in rebate_db.select_eq(area_ref, area).col("单位全名"):
        sales_this_client = sales_db.select_eq("单位全名", client_full_name)
        if not sales_this_client:
            continue
        kinds = sales_this_client.col("料型", unique=True)
        kinds = kind_order.sort(kinds)
        client_short_name = client_full_name.split()[-1]
        sheet.write_merge(row_idx, row_idx + len(kinds), 0, 0, client_short_name, XF_BOX_TC)

        # 写每月数据
        monthly_sales = defaultdict(int)
        for kind in kinds:
            sheet.write(row_idx, 1, kind, XF_BOX_C)
            sales_this_kind = sales_this_client.select_eq("料型", kind)
            annual_sales = 0
            for month in range(12):
                sales_this_month = sales_this_kind.select_fn("制单日期", lambda date: date[1] == month+1)
                if sales_this_month:
                    amount = sum(sales_this_month.col("销售数量"))
                    sheet.write(row_idx, month + 2, QUANT(amount), XF_BOX)
                    monthly_sales[month] += amount
                    global_monthly_sales[month] += amount
                    annual_sales += amount
                else:
                    sheet.write(row_idx, month + 2, '', XF_BOX)
            sheet.write(row_idx, 14, QUANT(annual_sales), XF_BOX)
            row_idx += 1

        # 写月度合计
        sheet.write(row_idx, 1, "合计", XF_BOX_C_RED)
        for month in range(12):
            amount = monthly_sales.get(month, "")
            sheet.write(row_idx, month + 2, QUANT(amount), XF_BOX_RED)
        sheet.write(row_idx, 14, QUANT(sum(monthly_sales.values())), XF_BOX_RED)
        row_idx += 1

    # 写年度合计
    for month in range(12):
        amount = global_monthly_sales.get(month, "")
        sheet.write(row_idx, month + 2, QUANT(amount), XF_BOX)
    sheet.write(row_idx, 14, QUANT(sum(global_monthly_sales.values())), XF_BOX)


def write_area_summary_sheet(sheet: Worksheet, sales_db: SalesDatabase):
    print("绘制区域汇总表")

    # 写header并调整列宽
    header = ["区域", "料型", *(f"{i + 1}月" for i in range(12)), "总计"]
    for i, text in enumerate(header):
        sheet.write(0, i, text, XF_BOX_C)
        sheet.col(i).width_mismatch = True
        sheet.col(i).width = 2048 if i < 1 else 1530

    # 写每个区域
    row_idx = 1
    for area in rebate_db.col(area_ref, unique=True):
        clients = set(rebate_db.select_eq(area_ref, area).col("单位全名"))
        sales_this_area = sales_db.select_fn("单位全名", lambda name: name in clients)
        kinds = sales_this_area.col("料型", unique=True)
        kinds = kind_order.sort(kinds)
        sheet.write_merge(row_idx, row_idx + len(kinds), 0, 0, area, XF_BOX_TC)

        # 写每月数据
        monthly_sales = defaultdict(int)
        for kind in kinds:
            sheet.write(row_idx, 1, kind, XF_BOX_C)
            sales_this_kind = sales_this_area.select_eq("料型", kind)
            annual_sales = 0
            for month in range(12):
                sales_this_month = sales_this_kind.select_fn("制单日期", lambda date: date[1] == month+1)
                if sales_this_month:
                    amount = sum(sales_this_month.col("销售数量"))
                    sheet.write(row_idx, month + 2, QUANT(amount), XF_BOX)
                    monthly_sales[month] += amount
                    annual_sales += amount
                else:
                    sheet.write(row_idx, month + 2, "", XF_BOX)
            sheet.write(row_idx, 14, QUANT(annual_sales), XF_BOX)
            row_idx += 1

        # 写月度合计
        sheet.write(row_idx, 1, "合计", XF_BOX_C_RED)
        for month in range(12):
            amount = monthly_sales.get(month, "")
            sheet.write(row_idx, month + 2, QUANT(amount), XF_BOX_RED)
        sheet.write(row_idx, 14, QUANT(sum(monthly_sales.values())), XF_BOX_RED)
        row_idx += 1


def write_area_summary_with_price_sheet(sheet: Worksheet, sales_db: SalesDatabase):
    print("绘制区域汇总表（费用系数）")

    # 写header并调整列宽
    header = ["区域", "料型", *(f"{m + 1}月" for m in range(12)), "总计"]
    for i, text in enumerate(header):
        sheet.write(0, i, text, XF_BOX_C)
        sheet.col(i).width_mismatch = True
        sheet.col(i).width = 2048 if i < 1 else 1530

    sheet.col(i + 1).width_mismatch = True
    sheet.col(i + 1).width = 580

    header = [*(f"{m + 1}月" for m in range(12)), "总计"]
    for i, text in enumerate(header, equiv_sales_i := i + 2):
        sheet.write(0, i, text, XF_BOX_C)
        sheet.col(i).width_mismatch = True
        sheet.col(i).width = 2048 if i < 1 else 1530

    sheet.col(i + 1).width_mismatch = True
    sheet.col(i + 1).width = 580

    for i, text in enumerate(header, equiv_price_i := i + 2):
        sheet.write(0, i, text, XF_BOX_C)
        sheet.col(i).width_mismatch = True
        sheet.col(i).width = 2048 if i < 1 else 1530

    # 写每个区域
    row_idx = 1
    for area in rebate_db.col(area_ref, unique=True):
        clients = set(rebate_db.select_eq(area_ref, area).col("单位全名"))
        sales_this_area = sales_db.select_fn("单位全名", lambda name: name in clients)
        kinds = sales_this_area.col("料型", unique=True)
        kinds = kind_order.sort(kinds)
        sheet.write_merge(row_idx, row_idx + len(kinds), 0, 0, area, XF_BOX_TC)

        # 写每月数据
        monthly_sales = defaultdict(int)
        monthly_equiv_sales = defaultdict(int)
        monthly_equiv_price = defaultdict(int)
        for kind in kinds:
            sheet.write(row_idx, 1, kind, XF_BOX_C)
            sales_this_kind = sales_this_area.select_eq("料型", kind)
            annual_sales = 0
            annual_equiv_sales = 0
            annual_equiv_price = 0

            for month in range(12):
                sales_this_month = sales_this_kind.select_fn("制单日期", lambda date: date[1] == month+1)
                if sales_this_month:
                    amount = sum(sales_this_month.col("销售数量"))
                    sheet.write(row_idx, month + 2, QUANT(amount), XF_BOX)
                    monthly_sales[month] += amount
                    annual_sales += amount

                    equiv_sales = amount * price_factor.factors.get(kind, 0)
                    sheet.write(row_idx, month + equiv_sales_i, QUANT(equiv_sales), XF_BOX)
                    monthly_equiv_sales[month] += equiv_sales
                    annual_equiv_sales += equiv_sales

                    equiv_price = equiv_sales * price_factor.prices.get(kind, 0)
                    sheet.write(row_idx, month + equiv_price_i, QUANT(equiv_price), XF_BOX)
                    monthly_equiv_price[month] += equiv_price
                    annual_equiv_price += equiv_price
                else:
                    sheet.write(row_idx, month + 2, "", XF_BOX)
                    sheet.write(row_idx, month + equiv_sales_i, "", XF_BOX)
                    sheet.write(row_idx, month + equiv_price_i, "", XF_BOX)

            sheet.write(row_idx, 14, QUANT(annual_sales), XF_BOX)
            sheet.write(row_idx, equiv_sales_i + 12, QUANT(annual_equiv_sales), XF_BOX)
            sheet.write(row_idx, equiv_price_i + 12, QUANT(annual_equiv_price), XF_BOX)
            row_idx += 1

        # 写月度合计
        sheet.write(row_idx, 1, "合计", XF_BOX_C_RED)
        for month in range(12):
            amount = monthly_sales.get(month, "")
            sheet.write(row_idx, month + 2, QUANT(amount), XF_BOX_RED)

            amount = monthly_equiv_sales.get(month, "")
            sheet.write(row_idx, month + equiv_sales_i, QUANT(amount), XF_BOX_RED)
            amount = monthly_equiv_price.get(month, "")
            sheet.write(row_idx, month + equiv_price_i, QUANT(amount), XF_BOX_RED)

        sheet.write(row_idx, 14, QUANT(sum(monthly_sales.values())), XF_BOX_RED)

        sheet.write(row_idx, equiv_sales_i + 12, QUANT(sum(monthly_equiv_sales.values())), XF_BOX_RED)
        sheet.write(row_idx, equiv_price_i + 12, QUANT(sum(monthly_equiv_price.values())), XF_BOX_RED)

        row_idx += 1


def convert_table(area_ref_: str, sales_db: SalesDatabase):
    global area_ref
    area_ref = area_ref_

    if area_ref == "湖大":  # 湖大只计入湖大料型
        sales_db = sales_db.select_fn("料型", lambda kind: kind_order.counted.get(kind, False))

    workbook = xlwt.Workbook(encoding="utf-8")
    for area in rebate_db.col(area_ref, unique=True):
        write_area_sheet(workbook.add_sheet(area), area, sales_db)
    
    if area_ref == "费用":
        write_area_summary_with_price_sheet(workbook.add_sheet("区域汇总"), sales_db)
    else:
        write_area_summary_sheet(workbook.add_sheet("区域汇总"), sales_db)

    path = f"销量透视表-{now_str}-{area_ref}.xls"    
    workbook.save(path)


def convert(sales_path: str, rebate_path: str, logger_callback: Callable[[str], None]):
    global rebate_db, kind_order, price_factor, area_ref, now_str
    now_str = datetime.now().strftime("%y%m%d-%H%M%S")

    logger_callback(f"(0/5) 打开：{sales_path}")
    book = xlrd.open_workbook(sales_path)
    sheet = book.sheet_by_index(0)
    sales_db = SalesDatabase(sheet)

    logger_callback(f"(1/5) 打开：{rebate_path}")
    rebate_db, kind_order, price_factor = read_rebate_book(rebate_path)

    logger_callback(f"(2/5) 转换中：湖大")
    convert_table("湖大", sales_db)

    logger_callback(f"(3/5) 转换中：中转站销量")
    convert_table("中转站销量", sales_db)

    logger_callback(f"(4/5) 转换中：费用")
    convert_table("费用", sales_db)

    logger_callback("(5/5) 转换成功！")


if __name__ == '__main__':
    convert(
        "25全年仓库 新表  明细数据.xls",
        "1. 2026年 新客户名-销量 返利 费用.xls",
        print,
    )
