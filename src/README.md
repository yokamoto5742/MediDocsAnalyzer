# 業務分析アプリ

## 概要
業務分析アプリは、日々の業務内容と時間を記録したWILLDOリスト(Excel)を読み込み、期間内の業務時間を集計して、テンプレートに沿ったExcelファイルに出力するWindows向けデスクトップツールです。

## ✨ 主な機能
- 年と期間(月・四半期・上半期・下半期・一年)を選ぶだけで、分析期間を自動入力
- 業務・デイリータスク・コミュニケーションを、クラーク業務などの観点から集計
- 「ファイルパス」ボタンで、WILLDOリスト・出力書式・出力先フォルダを画面から設定

## 動作要件
- Windows OS
- Python 3.13以上
- [uv](https://docs.astral.sh/uv/)
- 依存パッケージ(詳細は `pyproject.toml` を参照): openpyxl, polars, pyarrow, tkcalendar

## インストール方法
1. リポジトリをクローンまたはダウンロードします
2. プロジェクトフォルダで依存パッケージをインストールします
   ```bash
   uv sync
   ```
3. `utils/config.ini` の `[PATHS]` セクションを環境に合わせて設定します(起動後に「ファイルパス」ボタンからも設定できます)

## 使用方法
アプリケーションを起動します。
```bash
uv run python main.py
```

1. 「年」「年の始まり(1月始まり/4月始まり)」「期間」を選ぶと、開始日・終了日が自動で入ります(日付欄に `yyyy/mm/dd` 形式で直接入力することもできます)
2. 「分析開始」ボタンをクリックします
3. 結果が出力先フォルダに保存され、Excelで自動的に開きます

| ボタン | 動作 |
|---|---|
| 分析開始 | 入力された期間で分析し、結果のExcelを出力します。期間は次回起動時に復元されます |
| ファイルパス | WILLDOリスト・出力書式・出力先フォルダを選択して保存します |
| 設定 | `config.ini` をメモ帳で開きます |
| 閉じる | アプリを終了します |

## 入力ファイルの形式
- 各シートのA1セルに `2025年01月01日` 形式の日付が必要です(シート名は問いません)。「シート一覧」という名前のシートは読み込みません
- B列に業務内容、C列に時間(分)を入力します
- コミュニケーションは `内容(名前)` の形式で入力します
- 業務内容に「クラーク業務」を含むものは、「クラーク業務」として別に集計されます

## 設定ファイル
設定は `utils/config.ini` の1つだけです。ビルドした exe では、実行ファイルに同梱された `config.ini` を読み書きします。

| セクション | キー | 内容 |
|---|---|---|
| `[Appearance]` | `window_width` / `window_height` | ウィンドウの幅・高さ(px) |
| `[PATHS]` | `input_file_path` | WILLDOリスト(Excel)のパス |
| | `template_path` | 出力書式(テンプレートExcel)のパス |
| | `output_dir` | 分析結果の出力先フォルダ |
| `[Analysis]` | `start_date` / `end_date` | 前回の分析期間(アプリが自動で更新) |
| | `year_start_month` | 年の始まりの月(1 または 4)(アプリが自動で更新) |
| | `start_row` / `end_row` | 業務データの行範囲(両端を含む) |
| | `daily_task_start_row` / `daily_task_end_row` | デイリータスクの行範囲 |
| | `communication_start_row` / `communication_end_row` | コミュニケーションの行範囲 |

## 分析結果
結果は次の6シートを含むExcelファイルとして出力されます。

1. クラーク業務の集計
2. クラーク以外の業務の集計
3. デイリータスクの集計
4. コミュニケーションの集計(氏名と時間)
5. コミュニケーション内容の集計
6. 全項目の集計

出力ファイル名は `WILLDOリストまとめ{開始日}_{終了日}.xlsx` の形式です(例: `WILLDOリストまとめ20250101_20250331.xlsx`)。

## プロジェクト構成
```
TaskAnalyzer/
├── main.py                        # エントリーポイント
├── app/
│   ├── __init__.py                # バージョン情報(__version__)
│   ├── app_window.py              # メイン画面(TaskAnalyzerGUI)
│   ├── path_settings_dialog.py    # ファイルパス設定ダイアログ
│   └── constants.py               # 画面に表示するメッセージ
├── service/
│   ├── service_task_analyzer.py   # 分析全体の統括
│   ├── service_excel_reader.py    # WILLDOリスト(Excel)の読み込み
│   ├── service_data_analyzer.py   # Polars による集計
│   └── service_excel_writer.py    # テンプレートへの書き出し
├── utils/
│   ├── config.ini                 # 設定ファイル
│   ├── config_manager.py          # 設定ファイルの読み書き
│   └── date_period.py             # 年・期間から開始日・終了日を求める
├── assets/                        # アプリのアイコン
├── tests/                         # テスト(ソースと同じ構成)
└── docs/                          # CHANGELOG.md・LICENSE
```

## 開発者向け情報
### 開発コマンド
```bash
uv run pyright                       # 型チェック
uv run pytest tests/ -v --tb=short  # テスト(全件)
uv run python build.py               # exe のビルド(dist/ 配下に出力)
```

### コードの構成
- `TaskAnalyzerGUI`: GUIの作成と管理を担当
- `PathSettingsDialog`: ファイルパスの設定ダイアログ
- `TaskAnalyzer`: 分析処理全体の統括(保存先のパスを返し、エラーは GUI へ送出)
- `ExcelTaskReader`: Excelファイルからのデータ読み込み
- `TaskDataAnalyzer`: Polarsを使ったデータフレーム操作と集計
- `ExcelResultWriter`: 分析結果のExcel出力

### データフロー
1. `TaskAnalyzerGUI` がユーザー入力を受け取り、`TaskAnalyzer.run_analysis` を呼び出す
2. `TaskAnalyzer` が設定を読み込み、`ExcelTaskReader` でデータを読み込む
3. `TaskDataAnalyzer` が読み込んだデータを集計する
4. `ExcelResultWriter` が結果をテンプレートに書き出し、保存先のパスを返す

### 拡張方法
1. 新しい分析項目の追加
   - `service_data_analyzer.py` の `analyze_task_data` に新しい分析ロジックを追加し、戻り値のタプルに加える
   - テンプレートExcelに新しいシートを追加
   - `service_excel_writer.py` の `SHEET_NAMES` に、戻り値と同じ順で新しいシート名を追加

2. GUIの拡張
   - `app_window.py` の `_setup_gui` / `_setup_buttons` を修正
   - 画面に表示するメッセージは `app/constants.py` に定数として追加し、定数を参照する

3. データ処理の拡張
   - `service_excel_reader.py` を拡張して新しいデータの読み込み処理を追加

## 実装の詳細
- Polarsによるデータフレーム操作で、期間内の業務を高速に集計
- 期間の計算は `utils/date_period.py` で一元管理
- 正規表現で `内容(名前)` から名前を抽出
- 日付選択は tkcalendar の `DateEntry` を使用

## トラブルシューティング
1. アプリケーションが起動しない
   - Python 3.13以上がインストールされているか確認
   - `uv sync` を実行して必要なパッケージをインストール

2. 「指定された期間内のデータがありません。」と表示される
   - 指定した期間に該当するシートがあるか確認
   - 各シートのA1セルが `2025年01月01日` 形式の日付になっているか確認

3. ファイルが見つからない・保存できない
   - 「ファイルパス」ボタンで、WILLDOリスト・出力書式・出力先フォルダが正しく設定されているか確認
   - 同名の結果ファイルをExcelで開いたままだと保存できないため、閉じてから再実行

4. 分析結果が空または不完全
   - `config.ini` の行範囲(`start_row` など)が WILLDOリストの位置と合っているか確認
   - 業務内容(B列)と時間(C列)が入力されているか確認

## ライセンス
このプロジェクトのライセンス情報については、[LICENSE](docs/LICENSE) を参照してください。

## 更新履歴
更新履歴は [CHANGELOG.md](docs/CHANGELOG.md) を参照してください。
