#!/usr/bin/env python3
"""
配送記録シートのズレたデータを修復するスクリプト

使い方:
  python3 fix_delivery_records.py        → 状況を表示するだけ（プレビュー、安全）
  python3 fix_delivery_records.py --fix  → 実際に修復実行
"""
import sys
import json
import datetime
import re
import gspread
from google.oauth2.service_account import Credentials

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1VPO7xDMz_HPXyWuy8YVhHQQvmrmfGFG5bSuRDdtQ3DM"
CREDENTIALS_FILE = "credentials.json"
SHEET_NAME = "配送記録"
NUM_COLS = 9
DATE_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")

def col_letter(n):
    """列番号(1始まり)を A, B, ..., Z, AA, AB, ... に変換"""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def main():
    dry_run = "--fix" not in sys.argv

    print("🔌 Google Sheets接続中...")
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_url(SPREADSHEET_URL)
    sheet = spreadsheet.worksheet(SHEET_NAME)

    print(f"📊 シート '{SHEET_NAME}' を読み込み中...")
    all_values = sheet.get_all_values()
    if not all_values:
        print("シートが空です。終了。")
        return

    print(f"   行数: {len(all_values)}")
    max_col = max(len(row) for row in all_values)
    print(f"   最大列数: {max_col}")
    print()

    # バックアップ保存（常に取る）
    backup_file = f"配送記録_backup_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with open(backup_file, "w", encoding="utf-8") as f:
        json.dump(all_values, f, ensure_ascii=False, indent=2)
    print(f"💾 バックアップ作成: {backup_file}")
    print()

    rows = all_values[1:]  # ヘッダー除く

    bad_rows = []   # [(行番号, fixed_values, 列番号)]
    for i, row in enumerate(rows):
        row_num = i + 2
        date_col_idx = -1
        for col_idx, val in enumerate(row):
            if DATE_PATTERN.match(str(val).strip()):
                date_col_idx = col_idx
                break
        if date_col_idx == -1:
            continue  # 日付なしの行はスキップ

        # 日付位置から9列分取り出す
        extracted = row[date_col_idx : date_col_idx + NUM_COLS]
        while len(extracted) < NUM_COLS:
            extracted.append("")

        if date_col_idx != 0:
            bad_rows.append((row_num, extracted, date_col_idx + 1))

    print(f"📋 ズレた行: {len(bad_rows)}件")
    if bad_rows:
        print()
        print("  行番号 | 現在の列 | 日付       | 顧客コード | 名前")
        print("  ------ | -------- | ---------- | ---------- | --------")
        for row_num, vals, cur_col in bad_rows:
            date_v = vals[0]
            code_v = vals[2]
            name_v = vals[3]
            print(f"  {row_num:>5}  | 列{cur_col:>4}    | {date_v:<10} | {code_v:<10} | {name_v}")

    if not bad_rows:
        print("✅ ズレた行はありません。何もする必要なし。")
        return

    if dry_run:
        print()
        print("=" * 60)
        print("⚠️ これはプレビューです。実データは変更していません。")
        print()
        print("修復を実行するには：")
        print(f"   python3 {sys.argv[0]} --fix")
        print("=" * 60)
        return

    # ── 実際の修復処理 ─────────────────────────────────
    print()
    print("🔧 修復を実行します...")
    print()

    # ステップ1: バッチで A:I に正しい値を書き込み
    end_col_letter = col_letter(NUM_COLS)
    batch_data = []
    for row_num, vals, _ in bad_rows:
        batch_data.append({
            "range": f"A{row_num}:{end_col_letter}{row_num}",
            "values": [vals],
        })
    print(f"✏️ {len(batch_data)}件の行に正しいデータを書き込み中...")
    sheet.batch_update(batch_data)

    # ステップ2: J列以降のゴミデータをクリア
    if max_col > NUM_COLS:
        garbage_end = col_letter(max_col)
        clear_ranges = []
        for row_num, _, _ in bad_rows:
            clear_ranges.append(f"J{row_num}:{garbage_end}{row_num}")
        print(f"🧹 J列以降のゴミデータをクリア中...")
        sheet.batch_clear(clear_ranges)

    print()
    print(f"✅ 修復完了！{len(bad_rows)}件のズレを修正しました。")
    print(f"💾 バックアップ: {backup_file}")
    print()
    print("👉 Google Sheetsを開いて確認してください。")

if __name__ == "__main__":
    main()
