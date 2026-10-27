# テスト実施結果

- 実施日時: 2026年10月27日 14:55 JST
- 実行環境: Linux x86_64 のサンドボックス、Python 3.12.14
- 使用ライブラリ: pandas 3.1.0、openpyxl 3.1.5
- 結果: **依存パッケージのインストール成功、run_test.py 成功（終了コード0）**

## 実行したコマンド

リポジトリ直下で、専用の仮想環境を作成して実行しました。

```sh
python -m venv /workspace/work/cw-exp-sales-venv
/workspace/work/cw-exp-sales-venv/bin/python -m pip install --no-cache-dir -r cw-exp-sales/requirements.txt
/workspace/work/cw-exp-sales-venv/bin/python cw-exp-sales/run_test.py
/workspace/work/cw-exp-sales-venv/bin/python -m pip check
```

pip installは実際にパッケージを取得して成功しました。最初の試行は実行サンドボックスのネットワーク制限で失敗したため、ネットワークを許可したコマンドで再実行しました。pip checkも `No broken requirements found.` でした。

## 検証内容

- UTF-8 BOM付き・見出しあり2ファイル、CP932・見出しなし1ファイルを統合。
- 明細9件の全項目、数値型、売上金額、2シートの名前と列名をassertで確認。
- 店舗別合計: 草津店5,900円、大津店5,600円、彦根店3,100円。全店合計14,600円。
- 元CSV3ファイルのSHA-256が実行前後で一致。異常系の入力CSVについても一致。
- 別の作業フォルダーから引数なしで起動しても、スクリプトの隣のinput_csvを使用。
- 日本語や空白を含む入力先を、相対・絶対パスの両方で指定して検証。
- 空・空白のみ・見出しのみ・文字コード不明・引用符不正のファイルに日本語警告を出し、正常なファイルの処理を継続。
- 数字以外・空欄・NaN・無限大・掛け算のオーバーフローを警告してスキップ。列数不正・必須項目の空欄もスキップ。
- 小数と負数の計算、数式に見える商品名の文字列保存、CSV以外とサブフォルダーの除外。
- PermissionErrorを模擬し、日本語警告でスキップすることと、CSVを `r` モードで開くことを確認。
- 有効行なし・CSVなし・存在しないフォルダー・引数過多は終了コード1。既存Excelのハッシュは変化なし。

## 実行ログ

```text
売上CSVを確認しています: /workspace/codex-env/cw-exp-sales/input_csv
  01_sales_utf8.csv: 3 件を読み込みました。
  02_sales_utf8.csv: 3 件を読み込みました。
  03_sales_cp932.csv: 3 件を読み込みました。
集計が完了しました！ 明細 9 件、店舗 3 件です。
保存先: /workspace/codex-env/cw-exp-sales/output/売上集計.xlsx
元のCSVファイルは変更していません。
[PASS] サンプル3ファイル、9明細、店舗別合計、全店合計14,600円
[PASS] サンプルCSVのSHA-256は全3ファイルで実行前後一致
[PASS] UTF-8/CP932、見出し自動判定、相対/絶対パス、別の起動場所
[PASS] 空・読取不能・不正CSV、不正数値、列数不正でも有効な行を継続集計
[PASS] 小数・負数、文字列の保持、対象外ファイルとサブフォルダーの除外
[PASS] 異常系CSVもSHA-256が実行前後一致
[PASS] アクセス拒否は日本語警告でスキップし、CSVを読み取り専用で開く
[PASS] 有効行なし・CSVなし・フォルダーなし・引数過多: 終了コード1、既存Excelを維持
テストはすべて成功しました。元CSVは変更されていません。
サンプルExcel: /workspace/codex-env/cw-exp-sales/output/売上集計.xlsx
```

## サンプルCSVのSHA-256（実行前後で一致）

| ファイル | SHA-256 |
| --- | --- |
| 01_sales_utf8.csv | `e71672ad8bf647ea4690783d37c0957d0689b14b42165a1fb569d1a30260983d` |
| 02_sales_utf8.csv | `e4fc9ae524b8938a432da846268843439f9f11ae92134acd90190e36f0887cbd` |
| 03_sales_cp932.csv | `34ade38666c5615da4004da26d0171f42e66705bed1b4fdd8817e3055ea4e046` |

## 生成物と確認範囲

生成した `output/売上集計.xlsx` を同梱しています。テスト用の異常なCSVは一時フォルダーだけに作成し、input_csvには正常な3ファイルだけを残しています。

集計処理はLinuxサンドボックスで実行しました。Windows 11の実機、およびcmd.exe上でのrun.batのダブルクリック実行はこの環境では未検証です。run.batはUTF-8（BOMなし）・CRLFで保存し、chcp 65001とPYTHONUTF8=1を設定しています。
