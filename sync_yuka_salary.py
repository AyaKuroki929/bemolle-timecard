#!/usr/bin/env python3
"""
タイムカード → 有加報酬.numbers 日当自動入力スクリプト
「6月分作って」で呼び出す。日当（B列）を自動入力し、施術料（C〜G列）は手入力のまま。

使い方:
  python3 sync_yuka_salary.py          # 先月分
  python3 sync_yuka_salary.py 2026-06  # 指定月
"""

import sys
import warnings
warnings.filterwarnings("ignore")

import requests
import json
from datetime import datetime, date
import numbers_parser
from numbers_parser import Document
from numbers_parser.cell import NumberCell, EmptyCell
import subprocess
import os

# ── 設定 ──────────────────────────────────────────────────────────────
GAS_URL    = "https://script.google.com/macros/s/AKfycbw9SmXqYA0i_Vysp4l0ZbRcNjvTs0miOm-CFVcF9kWLGm-VZ43sqo9FWT0rpgf0PJq4kA/exec"
NUMBERS_FILE = "/Users/ayakuroki/Library/Mobile Documents/com~apple~CloudDocs/紫妃彩/Bemolle/スタッフ/業務委託/有加報酬.numbers"
STAFF_NAME = "中田有加"
# ──────────────────────────────────────────────────────────────────────

def get_month_key(arg=None):
    """引数なし→先月、YYYY-MM形式→そのまま"""
    if arg:
        return arg  # e.g. "2026-06"
    now = datetime.now()
    if now.month == 1:
        return f"{now.year - 1}-12"
    return f"{now.year}-{now.month - 1:02d}"

def fetch_allowances(month_key):
    """GAS APIから日当データ取得"""
    year, month = month_key.split("-")
    from_date = f"{year}-{month}-01"
    # 月末計算
    import calendar
    last_day = calendar.monthrange(int(year), int(month))[1]
    to_date = f"{year}-{month}-{last_day:02d}"

    res = requests.get(GAS_URL, params={
        "action": "getDailyAllowance",
        "from": from_date,
        "to": to_date,
    }, timeout=30)
    data = res.json()
    if not data.get("ok"):
        raise RuntimeError(f"GAS error: {data.get('error')}")
    return data.get("data", [])

def month_label(month_key):
    """2026-06 → 2026年6月"""
    year, month = month_key.split("-")
    return f"{year}年{int(month)}月"

def make_sheet(month_key, allowances):
    """Numbers ファイルに新しい月シートを作成して日当を入力"""

    label = month_label(month_key)

    # 日当があった日だけ抽出（日付 → 金額のdict）
    daily = {}
    for rec in allowances:
        if rec.get("staff") == STAFF_NAME and rec.get("amount"):
            try:
                amt = float(str(rec["amount"]).replace("¥", "").replace(",", ""))
                if amt > 0:
                    day = int(rec["date"].split("-")[2])
                    daily[day] = amt
            except (ValueError, IndexError):
                pass

    if not daily:
        print(f"⚠️  {label} の日当データがありません（打刻・メモなし出勤が0件）")
        return

    print(f"📅 {label}：{len(daily)} 日出勤, 日当合計 ¥{sum(daily.values()):,.0f}")

    # ── Numbers ファイルを開く ──────────────────────────────────────
    doc = Document(NUMBERS_FILE)
    sheets = doc.sheets

    # 既存シート名チェック
    existing_names = [s.name for s in sheets]
    if label in existing_names:
        print(f"✅ '{label}' シートはすでに存在します。日当列だけ上書きします。")
        sheet = next(s for s in sheets if s.name == label)
        table = sheet.tables[0]
        _fill_allowance_column(table, daily)
    else:
        # 最後のシートをコピーして新しいシートを作る
        # numbers-parser はシートコピー未対応なので、前月シートの構造を参考に新規作成
        sheet = doc.add_sheet(label)
        table = sheet.tables[0]
        _build_table(table, label, daily)
        print(f"✅ '{label}' シートを作成しました")

    doc.save(NUMBERS_FILE)
    print(f"💾 保存しました: {NUMBERS_FILE}")

def _apply_currency_format(cell):
    """¥カンマ書式を適用（_num_format_id=2 が必須）"""
    try:
        cell.set_cell_formatting(
            "currency",
            currency_code="JPY",
            decimal_places=0,
            show_thousands_separator=True,
        )
        cell._num_format_id = 2
    except Exception:
        pass

def _build_table(table, label, daily):
    """26行×8列のテーブルを構築"""
    # ── 行・列数を調整 ──
    while table.num_rows < 26:
        table.append_row([])
    while table.num_cols < 8:
        table.append_column([])

    # ── ヘッダー行（Row 0） ──
    table[0][0].value = "日付"
    table[0][7].value = "合計"

    # ── データ行（Row 1〜24）：日当のある日だけ入力 ──
    row_idx = 1
    for day in sorted(daily.keys()):
        if row_idx > 24:
            break
        amt = daily[day]
        table[row_idx][0].value = float(day)   # A列: 日付
        table[row_idx][1].value = amt           # B列: 日当
        _apply_currency_format(table[row_idx][1])
        # H列: SUM(B:G) 数式
        row_num = row_idx + 1  # Numbers は1始まり
        table[row_idx][7].formula = f"SUM(B{row_num}:G{row_num})"
        row_idx += 1

    # ── フッター行（Row 25） ──
    table[25][0].value = "合計"
    table[25][1].value = f"{len(daily)} 日出勤"
    table[25][7].formula = "SUM(H2:H25)"
    _apply_currency_format(table[25][7])

def _fill_allowance_column(table, daily):
    """既存シートのB列（日当）を上書き"""
    for row_idx in range(1, 25):
        try:
            day_cell = table[row_idx][0]
            if day_cell.value is None:
                continue
            day = int(float(str(day_cell.value)))
            if day in daily:
                table[row_idx][1].value = daily[day]
                _apply_currency_format(table[row_idx][1])
        except (ValueError, TypeError, IndexError):
            continue

def main():
    arg = sys.argv[1] if len(sys.argv) > 1 else None
    month_key = get_month_key(arg)
    print(f"🔄 {month_label(month_key)} の日当を取得中...")

    allowances = fetch_allowances(month_key)
    make_sheet(month_key, allowances)

if __name__ == "__main__":
    main()
