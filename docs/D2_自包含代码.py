# -*- coding: utf-8 -*-
"""D2 三态分流 · 自包含代码（直接读写 Excel，无需影刀变量映射）

用法：影刀新建 PC 自动化应用 → 自定义搭建 → 拖一条「插入代码段(Python)」
      → 把本文件全部内容粘贴到代码区 → 运行。

依赖：openpyxl。若影刀 Python 环境未装，代码会自动尝试 pip 安装；
      若自动安装失败，按报错提示在影刀菜单 → Python → pip 手动安装 openpyxl。
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

SRC = r"C:\Users\jjchen\WorkBuddy\2026-11-02-18-03-51\rpa-order-audit\data\模拟订单.xlsx"
DST = r"C:\Users\jjchen\WorkBuddy\2026-11-02-18-03-51\rpa-order-audit\data\审核结果_D2.xlsx"


def is_11digit(v):
    if v is None:
        return False
    s = str(v).strip()
    if isinstance(v, float) and v.is_integer():
        s = str(int(v))
    return len(s) == 11 and s.isdigit()


def judge(row):
    (oid, plat, buyer, phone, prov, city, dist, addr, goods,
     qty, price, cost, payable, paid, stock, returns, paystatus, note) = row[:18]
    note = note or ""
    hits = []
    if paid == 0 and paystatus == "已支付":
        hits.append("R1")
    if paid < cost * qty:
        hits.append("R2")
    if qty > stock:
        hits.append("R3")
    if not is_11digit(phone):
        hits.append("M1")
    if not prov or not city or not dist or not addr:
        hits.append("M2")
    if any(k in str(note) for k in ("转寄", "代收", "自提")):
        hits.append("M3")
    if (returns or 0) >= 3:
        hits.append("M4")
    hard = [h for h in hits if h.startswith("R")]
    soft = [h for h in hits if h.startswith("M")]
    if hard:
        return "自动拦截", ",".join(hard + soft)
    if soft:
        return "转人工", ",".join(soft)
    return "自动通过", ""


wb = load_workbook(SRC, data_only=True)
ws = wb.active

rows = list(ws.iter_rows(values_only=True))
out_rows = [list(rows[0]) + ["判定", "命中规则"]]


for r in rows[1:]:
    verdict, hits = judge(r)
    out_rows.append(list(r) + [verdict, hits])

out = Workbook()
ows = out.active
ows.title = "审核结果"
for row in out_rows:
    ows.append(row)
for c in ows[1]:
    c.font = Font(bold=True)
out.save(DST)

n_pass = sum(1 for r in out_rows[1:] if r[-2] == "自动通过")
n_manual = sum(1 for r in out_rows[1:] if r[-2] == "转人工")
n_block = sum(1 for r in out_rows[1:] if r[-2] == "自动拦截")
print("D2 完成: 自动通过 %d / 转人工 %d / 自动拦截 %d -> %s" % (n_pass, n_manual, n_block, DST))
