"""
振宇班次查询脚本
表格结构：
  - C198 = 2026/2/2（周一），D198=2/3...I198=2/8（周日）
  - 每周块纵向占 7 行，下一周日期行 = 当前日期行 + 7
  - B 列为班次标签（早班/中班/晚班/备班等）
"""

import os
import sys
import time
import hmac
import hashlib
import base64
import urllib.parse
from datetime import datetime, timedelta, timezone
import requests
from playwright.sync_api import sync_playwright

DOC_URL = "https://docs.qq.com/sheet/DUEVHc3N4am1iZ1ZG?tab=BB08J2"
PERSON = "振宇"
DINGTALK_WEBHOOK = os.environ.get("DINGTALK_WEBHOOK", "")
DINGTALK_* = os.environ.get("DINGTALK_*", "")

ANCHOR_DATE = datetime(2026, 2, 2)
ANCHOR_ROW = 198
WEEK_ROWS = 7
COL_LETTERS = ["C", "D", "E", "F", "G", "H", "I"]  # 周一到周日


def get_date_cell(target: datetime):
    delta = (target.date() - ANCHOR_DATE.date()).days
    week_offset = delta // 7
    day_offset = delta % 7
    row = ANCHOR_ROW + week_offset * WEEK_ROWS
    col = COL_LETTERS[day_offset]
    return row, col


def read_cell(page, cell_ref: str) -> str:
    name_box = page.get_by_role("textbox").first
    name_box.fill(cell_ref)
    name_box.press("Enter")
    page.wait_for_timeout(400)
    formula_bar = page.get_by_role("combobox").first
    return formula_bar.inner_text().strip()


def find_shift(page, target_date: datetime):
    row, col = get_date_cell(target_date)
    date_str_variants = [
        f"{target_date.year}/{target_date.month}/{target_date.day}",
        f"{target_date.year}-{target_date.month:02d}-{target_date.day:02d}",
    ]
    print(f"目标日期: {target_date.strftime('%Y-%m-%d')}，预计单元格: {col}{row}")

    # 扩大扫描范围到 ±8 行，帮助诊断偏移
    date_row = None
    scan_range = list(range(row - 8, row + 9))
    for r in scan_range:
        val = read_cell(page, f"{col}{r}")
        print(f"  {col}{r} = {val!r}")
        if any(v in val for v in date_str_variants):
            date_row = r
            print(f"  ✓ 命中日期行 {r}，偏移={r - row}")
            break

    if date_row is None:
        print("  ✗ 扫描范围内未找到日期行，打印 B 列参考：")
        for r in range(row - 4, row + 5):
            val = read_cell(page, f"B{r}")
            print(f"  B{r} = {val!r}")
        return None

    for offset in range(1, 8):
        r = date_row + offset
        val = read_cell(page, f"{col}{r}")
        print(f"  {col}{r} = {val!r}")
        if PERSON in val:
            shift = read_cell(page, f"B{r}")
            print(f"  ✓ 振宇在第{r}行，B{r} = {shift!r}")
            return shift

    print("  ✗ 未找到振宇")
    return None


def build_dingtalk_url() -> str:
    if not DINGTALK_*:
        return DINGTALK_WEBHOOK
    timestamp = str(round(time.time() * 1000))
    string_to_sign = timestamp + "\n" + DINGTALK_*
    hmac_code = hmac.new(
        DINGTALK_*.encode("utf-8"),
        string_to_sign.encode("utf-8"),
        digestmod=hashlib.sha256,
    ).digest()
    sign = urllib.parse.quote_plus(base64.b64encode(hmac_code))
    return f"{DINGTALK_WEBHOOK}&timestamp={timestamp}&sign={sign}"


def notify(shift: str, target_date: datetime):
    if not DINGTALK_WEBHOOK:
        print("ERROR: 未配置 DINGTALK_WEBHOOK 环境变量")
        sys.exit(1)
    weekdays = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
    text = (
        f"明天（{target_date.month}月{target_date.day}日 "
        f"{weekdays[target_date.weekday()]}）振宇上 **{shift}**"
    )
    url = build_dingtalk_url()
    resp = requests.post(
        url,
        json={"msgtype": "markdown", "markdown": {"title": "振宇明天班次", "text": text}},
        timeout=10,
    )
    res = resp.json()
    if res.get("errcode") == 0:
        print(f"✓ 已发送: {text}")
    else:
        print(f"ERROR 发送失败: {res}")
        sys.exit(1)


def main():
    cst = timezone(timedelta(hours=8))
    tomorrow = datetime.now(cst) + timedelta(days=1)
    tomorrow = tomorrow.replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=None)
    print(f"查询明天班次: {tomorrow.strftime('%Y-%m-%d')}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(viewport={"width": 1600, "height": 900})
        page = ctx.new_page()
        page.goto(DOC_URL, wait_until="load", timeout=90000)
        page.wait_for_timeout(5000)
        shift = find_shift(page, tomorrow)
        browser.close()

    if shift is None:
        print("未找到振宇班次，跳过通知")
        return

    shift = shift.strip()
    if "备班" in shift:
        print("明天备班（休息），不提醒")
        return

    notify(shift, tomorrow)


if __name__ == "__main__":
    main()
