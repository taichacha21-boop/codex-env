"""店舗の売上CSVをまとめて、明細と店舗別合計をExcelに保存します。"""

import csv
import math
import os
from pathlib import Path
import sys


# 起動した場所ではなく、このスクリプトの場所を基準にします。
BASE_DIR = Path(__file__).resolve().parent
COLUMNS = ["日付", "店舗名", "商品名", "数量", "単価"]
OUTPUT_PATH = BASE_DIR / "output" / "売上集計.xlsx"


def read_sales(csv_path):
    """元のCSVは読み取り専用で開き、有効な行だけを返します。"""
    rows = None
    for encoding in ("utf-8-sig", "cp932"):
        try:
            # ファイル全体を読んでから処理し、文字コードの再試行時の重複を防ぎます。
            with csv_path.open("r", encoding=encoding, newline="") as source:
                reader = csv.reader(source, strict=True)
                rows = [(reader.line_num, row) for row in reader]
            break
        except UnicodeDecodeError:
            continue
        except (OSError, csv.Error):
            print(f"【警告】{csv_path.name} は読み取れないため、スキップしました。")
            return []

    if rows is None:
        print(f"【警告】{csv_path.name} の文字コードを読み取れません。スキップしました。")
        return []

    sales = []
    first_row = True
    has_data = False
    for line_number, row in rows:
        values = [value.strip() for value in row]
        if not any(values):
            continue
        # 先頭の空行を除き、最初の行が列名と一致した場合だけ見出しとします。
        if first_row:
            first_row = False
            if values == COLUMNS:
                continue
        has_data = True
        location = f"{csv_path.name} の {line_number} 行目"
        if len(values) != len(COLUMNS):
            print(f"【警告】{location} は列数が5列ではないため、スキップしました。")
            continue
        if not all(values[:3]):
            print(f"【警告】{location} は日付・店舗名・商品名に空欄があるため、スキップしました。")
            continue
        try:
            quantity = float(values[3])
            unit_price = float(values[4])
            amount = quantity * unit_price
            # 空欄・文字・NaN・無限大などは集計に含めません。
            if not all(math.isfinite(number) for number in (quantity, unit_price, amount)):
                raise ValueError
        except ValueError:
            print(f"【警告】{location} は数量または単価が正しい数値ではないため、スキップしました。")
            continue
        sales.append(values[:3] + [quantity, unit_price, amount])

    if not has_data:
        print(f"【警告】{csv_path.name} は空、または見出しだけのため、スキップしました。")
    return sales


def main():
    if len(sys.argv) > 2:
        print('使い方: python main.py "CSVフォルダーのパス"（省略できます）')
        return 1

    input_dir = Path(sys.argv[1]) if len(sys.argv) == 2 else Path("input_csv")
    if not input_dir.is_absolute():
        input_dir = BASE_DIR / input_dir
    if not input_dir.is_dir():
        print(f"CSVフォルダーが見つかりません: {input_dir}")
        print("フォルダーを作り、売上CSVを入れてから、もう一度実行してください。")
        return 1

    try:
        import pandas as pd
    except ImportError:
        print("必要な部品が見つかりません。次のコマンドを実行してください。")
        print('python -m pip install -r requirements.txt')
        return 1

    print(f"売上CSVを確認しています: {input_dir}")
    try:
        csv_files = sorted(path for path in input_dir.glob("*.csv") if path.is_file())
    except OSError:
        print("CSVフォルダーを開けません。場所やアクセス権を確認してください。")
        return 1
    if not csv_files:
        print("CSVファイルがありません。フォルダーに .csv ファイルを入れてください。")
        return 1

    sales = []
    for csv_path in csv_files:
        rows = read_sales(csv_path)
        sales.extend(rows)
        print(f"  {csv_path.name}: {len(rows)} 件を読み込みました。")
    if not sales:
        print("集計できるデータがありません。Excelは作成・更新していません。")
        return 1

    details = pd.DataFrame(sales, columns=COLUMNS + ["売上金額"])
    totals = details.groupby("店舗名", as_index=False)["売上金額"].sum()
    totals = totals.rename(columns={"売上金額": "売上合計"})
    try:
        OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
            details.to_excel(writer, sheet_name="明細", index=False)
            totals.to_excel(writer, sheet_name="店舗別合計", index=False)
            # 日付や名前を、そのまま文字として保存します（「=」で始まる名前も含む）。
            for sheet_name, text_columns in (("明細", 3), ("店舗別合計", 1)):
                sheet = writer.sheets[sheet_name]
                for cells in sheet.iter_rows(min_row=2, max_col=text_columns):
                    for cell in cells:
                        cell.data_type = "s"
    except ImportError:
        print("Excelを保存する部品が見つかりません。次のコマンドを実行してください。")
        print('python -m pip install -r requirements.txt')
        return 1
    except PermissionError:
        print("Excelに保存できませんでした。売上集計.xlsx を閉じ、保存先のアクセス権も確認してください。")
        return 1
    except (OSError, ValueError) as error:
        print(f"Excelに保存できませんでした。保存先や空き容量を確認してください。詳細: {error}")
        return 1

    print(f"集計が完了しました！ 明細 {len(details)} 件、店舗 {len(totals)} 件です。")
    print(f"保存先: {OUTPUT_PATH}")
    print("元のCSVファイルは変更していません。")
    return 0


if __name__ == "__main__":
    exit_code = main()
    # run.bat から引数なしで起動したときだけ、結果を確認できるように待ちます。
    if len(sys.argv) == 1 and os.environ.get("CW_EXP_SALES_PAUSE") == "1":
        try:
            input("Enterキーを押すと終了します。")
        except (EOFError, KeyboardInterrupt):
            pass
    sys.exit(exit_code)
