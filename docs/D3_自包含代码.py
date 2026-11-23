# -*- coding: utf-8 -*-
"""D3 对账 · 自包含代码（直接读写 Excel，无需影刀变量映射）

用法：影刀新建 PC 自动化应用 → 自定义搭建 → 拖一条「插入代码段(Python)」
      → 把本文件全部内容粘贴到代码区 → 运行。

对账口径：应结金额 = 实付金额 − 平台扣点 − 退款 + 补差
         账单金额 == 应结金额 → 平；否则记为差异并标类型。
"""
import sys
import subprocess
from rich import print as rprint

def _ensure(pkg):
    try:
        __import__(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

_ensure("openpyxl")
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font

ORDER_FILE = r"C:\Users\jjchen\WorkBuddy\2026-11-02-18-03-51\rpa-order-audit\data\模拟订单.xlsx"
BILL_FILE = r"C:\Users\jjchen\WorkBuddy\2026-11-02-18-03-51\rpa-order-audit\data\平台账单.xlsx"
DST = r"C:\Users\jjchen\WorkBuddy\2026-11-02-18-03-51\rpa-order-audit\data\对账结果.xlsx"


def read_table(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    header = list(rows[0])
    data = []
    for r in rows[1:]:
        data.append(dict(zip(header, r)))
    return header, data


def money(x):
    if x is None:
        return 0.0
    return round(float(x), 2)


# 1. 读两张表
_, orders = read_table(ORDER_FILE)
_, bills = read_table(BILL_FILE)


order_map = {str(o["订单号"]).strip(): o for o in orders}
bill_map = {str(b["订单号"]).strip(): b for b in bills}


# 2. 按订单号全量匹配（并集，保证缺单也能被找出）
all_ids = sorted(set(order_map) | set(bill_map))

out = []
for oid in all_ids:
    o = order_map.get(oid)
    b = bill_map.get(oid)
    platform = (o or b).get("平台")

    if o and not b:
        out.append([oid, platform, money(o["实付金额"]), None, None, "账单缺单", None,
                    "订单表有、账单表无，平台未结算"])
        continue
    if b and not o:
        out.append([oid, platform, None, money(b["账单金额"]), None, "订单缺单", None,
                    "账单表有、订单表无，疑似非本系统订单"])
        continue

    paid = money(o["实付金额"])
    bill_amt = money(b["账单金额"])
    fee = money(b["平台扣点"])
    refund = money(b["退款"])
    extra = money(b["补差"])
    should = round(paid - fee - refund + extra, 2)
    diff = round(bill_amt - should, 2)
    if abs(diff) < 0.005:
        out.append([oid, platform, paid, bill_amt, should, "平", 0.0, ""])
    else:
        out.append([oid, platform, paid, bill_amt, should, "金额不符", diff,
                    "账单金额与应结金额差 %+.2f 元" % diff])

# 3. 写出结果
wb = Workbook()
ws = wb.active
ws.title = "对账结果"
header = ["订单号", "平台", "实付金额", "账单金额", "应结金额", "对账结果", "差额", "说明"]
ws.append(header)
for row in out:
    ws.append(row)
for c in ws[1]:
    c.font = Font(bold=True)
wb.save(DST)

from collections import Counter
cnt = Counter(r[5] for r in out)
print("D3 完成: 平 %d / 金额不符 %d / 账单缺单 %d / 订单缺单 %d -> %s" % (
    cnt.get("平", 0), cnt.get("金额不符", 0),
    cnt.get("账单缺单", 0), cnt.get("订单缺单", 0), DST))
