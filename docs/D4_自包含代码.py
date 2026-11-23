# -*- coding: utf-8 -*-
"""D4 告警汇总 + 指纹幂等 · 自包含代码（直接读写 Excel，无需影刀变量映射）

用法：影刀新建 PC 自动化应用 → 自定义搭建 → 拖一条「插入代码段(Python)」
      → 把本文件全部内容粘贴到代码区 → 运行。

功能：
  1. 汇总异常：读 D2 审核结果（转人工/自动拦截）+ D3 对账结果（三类差异）；
  2. 生成一条人能读懂的飞书告警消息（分类、订单级、可追溯）；
  3. 指纹幂等：按"日期 + 异常内容"算 md5 存标记文件，重跑内容不变则跳过推送，
     防止同批异常反复轰炸；异常清单变化（新单/改判）则自动重发；
  4. 飞书推送：复用项目一 .env 的 FEISHU_WEBHOOK，四态结果 + 退避重试，
     失败不抛异常、不崩主流程，落盘一份可人工补发。

开关注释（在下方 CONFIG 区）：
  SEND_REAL = False  # 默认演练：只落文件、打印消息，不真发飞书
  想真发时改成 True（会读取项目一 .env 里的 webhook）。
"""
import sys
import os
import json
import time
import hashlib
import subprocess
from datetime import datetime


def _ensure(pkg):
    try:
        __import__(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])


_ensure("openpyxl")
from openpyxl import load_workbook

# ======================= CONFIG（按需改） =======================
BASE = r"C:\Users\jjchen\WorkBuddy\2026-11-02-18-03-51"
D2_FILE = BASE + r"\rpa-order-audit\data\审核结果_D2.xlsx"
D3_FILE = BASE + r"\rpa-order-audit\data\对账结果.xlsx"
STATE_FILE = BASE + r"\rpa-order-audit\data\alert_state.json"
MSG_FILE = BASE + r"\rpa-order-audit\data\告警消息.txt"
ENV_FILE = BASE + r"\ecom-daily-report\.env"   # 复用项目一的 webhook
SEND_REAL = False          # False=演练(不真发)；True=真发到飞书
# ==============================================================


def read_table(path):
    """读 Excel 首个工作表，返回 dict 列表（表头做 key）。"""
    wb = load_workbook(path, data_only=True)
    rows = list(wb.active.iter_rows(values_only=True))
    header = list(rows[0])
    return [dict(zip(header, r)) for r in rows[1:]]


def load_webhook():
    """从项目一 .env 读 FEISHU_WEBHOOK，不打印、不硬编码。找不到返回 None。"""
    if not os.path.exists(ENV_FILE):
        return None
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line.startswith("FEISHU_WEBHOOK="):
                v = line.split("=", 1)[1].strip().strip('"').strip("'")
                return v or None
    return None


# ---------- 1. 汇总异常（订单级，可追溯） ----------
audit = read_table(D2_FILE)
recon = read_table(D3_FILE)

blocked = [r for r in audit if r.get("判定") == "自动拦截"]
manual = [r for r in audit if r.get("判定") == "转人工"]
diffs = [r for r in recon if r.get("对账结果") not in (None, "", "平")]

# 指纹素材：只取"会改变告警语义"的字段，排序后拼接，避免行序抖动导致误判为新内容
finger_items = []
for r in blocked:
    finger_items.append("拦截|%s|%s" % (r["订单号"], r.get("命中规则") or ""))
for r in manual:
    finger_items.append("人工|%s|%s" % (r["订单号"], r.get("命中规则") or ""))
for r in diffs:
    finger_items.append("对账|%s|%s|%s" % (r["订单号"], r["对账结果"], r.get("差额")))
finger_core = "|".join(sorted(finger_items))
today = datetime.now().strftime("%Y-%m-%d")
# 模式纳入指纹：演练(dry)与真发(real)互不干扰——从演练切真发时指纹变化，能正常补发
mode = "real" if SEND_REAL else "dry"
fingerprint = hashlib.md5((today + "||" + mode + "||" + finger_core).encode("utf-8")).hexdigest()


# ---------- 2. 生成告警消息 ----------
lines = [
    "【订单审核日报】%s" % today,
    "共发现异常 %d 单（拦截 %d / 转人工 %d / 对账差异 %d）" % (
        len(blocked) + len(manual) + len(diffs), len(blocked), len(manual), len(diffs)),
    "",
]
if blocked:
    lines.append("■ 自动拦截 %d 单（已冻结，禁止发货）" % len(blocked))
    for r in blocked:
        lines.append("  - %s %s 命中%s" % (r["订单号"], r.get("商品"), r.get("命中规则")))
    lines.append("")
if manual:
    lines.append("■ 转人工 %d 单（进人工队列核对）" % len(manual))
    for r in manual:
        lines.append("  - %s %s 命中%s" % (r["订单号"], r.get("商品"), r.get("命中规则")))
    lines.append("")
if diffs:
    lines.append("■ 对账差异 %d 单（需财务核对）" % len(diffs))
    for r in diffs:
        d = r.get("差额")
        amt = ("%+.2f元" % float(d)) if d not in (None, "") else "—"
        lines.append("  - %s %s %s" % (r["订单号"], r["对账结果"], amt))
    lines.append("")
lines.append("生成时间 %s" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
msg = "\n".join(lines)

# 落盘一份（存档是证据，一条不丢；也可人工补发）
with open(MSG_FILE, "w", encoding="utf-8") as f:
    f.write(msg)


# ---------- 3. 指纹幂等判断 ----------
prev = None
if os.path.exists(STATE_FILE):
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            prev = json.load(f)
    except Exception:
        prev = None

if prev and prev.get("fingerprint") == fingerprint and prev.get("status") in ("ok", "dry_run"):
    print("D4 幂等跳过：异常内容与上次一致（指纹 %s），不重复推送。" % fingerprint[:8])
    print("上次处理时间：%s，上次结果：%s" % (prev.get("sent_at"), prev.get("status")))
    # 幂等跳过也要保证消息文件在（首次已落盘；若被误删则补写）
    if not os.path.exists(MSG_FILE):
        with open(MSG_FILE, "w", encoding="utf-8") as f:
            f.write(msg)
    sys.exit(0)


# ---------- 4. 飞书推送（四态 + 退避重试，失败不崩主流程） ----------
def send_feishu(text, webhook):
    """返回终态字符串：ok / failed / skipped_no_webhook / dry_run。永不抛异常。"""
    if not SEND_REAL:
        return "dry_run"
    if not webhook:
        return "skipped_no_webhook"
    import urllib.request
    payload = json.dumps({"msg_type": "text", "content": {"text": text}},
                         ensure_ascii=False).encode("utf-8")
    for attempt in range(1, 4):
        if attempt > 1:
            time.sleep((1, 3)[attempt - 2])
        try:
            req = urllib.request.Request(
                webhook, data=payload, headers={"Content-Type": "application/json"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                body = json.loads(resp.read().decode("utf-8"))
            if body.get("code", 0) == 0 and body.get("status_code", 0) == 0:
                return "ok"
            # 飞书业务码非 0：配置/内容问题，重试无意义
            print("飞书返回业务错误（不重试）:", body)
            return "failed"
        except Exception as exc:
            print("飞书推送第 %d 次失败: %s" % (attempt, exc))
    return "failed"


webhook = load_webhook()
status = send_feishu(msg, webhook)

# 发送成功或演练完成都写指纹（记住"这批已处理"）；真发失败不写，下次可补发
if status in ("ok", "dry_run"):
    state = {"fingerprint": fingerprint, "status": status,
             "sent_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
             "counts": {"拦截": len(blocked), "转人工": len(manual), "对账差异": len(diffs)}}
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

print("=" * 50)
print(msg)
print("=" * 50)
print("推送结果: %s | 指纹: %s | 消息已存: %s" % (status, fingerprint[:8], MSG_FILE))
if status == "dry_run":
    print("（演练模式：未真发。想真发把 SEND_REAL 改为 True 重跑。）")
