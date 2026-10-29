# 医療文書集計アプリ

## 概要
医療文書集計アプリは、医療文書担当一覧(Excel)を医療文書担当一覧データベースに取り込み、指定期間内の文書作成件数を担当者別・診療科別に集計して、テンプレートに沿ったExcelファイルに出力するWindows向けデスクトップツールです。

## ✨ 主な機能
- 元データをデータベースに取り込み、重複を除いて保存(取り込み後に自動でバックアップ)
- 年と期間(月・四半期・上半期・下半期・一年)を選ぶだけで、分析期間を自動入力(1月始まり/4月始まりを選択可能)
- 担当者×診療科の文書作成件数を集計してExcelに出力
- 「ファイルパス」ボタンで、元データ・データベース・バックアップ先・出力書式・出力先を画面から設定

## 動作要件
- Windows OS
- Python 3.13以上
- [uv](https://docs.astral.sh/uv/)
- 依存パッケージ(詳細は `pyproject.toml` を参照): babel, openpyxl, polars, pyarrow, tkcalendar

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

1. 「データ読込」ボタンで、元データをデータベースに取り込みます
2. 「年」「年の始まり(1月始まり/4月始まり)」「期間」を選ぶと、開始日・終了日が自動で入ります(日付欄に `yyyy/mm/dd` 形式で直接入力することもできます)
3. 「集計開始」ボタンをクリックします
4. 結果が出力先フォルダに保存され、Excelで自動的に開きます

| ボタン | 動作 |
|---|---|
| データ読込 | 元データをデータベースに取り込み、バックアップを作成します |
| 集計開始 | 入力された期間で集計し、結果のExcelを出力します。期間は次回起動時に復元されます |
| ファイルパス | 元データ・データベース・バックアップ先・出力書式・出力先を選択して保存します |
| 設定 | `config.ini` をメモ帳で開きます |
| 閉じる | アプリを終了します |

## 入力ファイルの形式
- 1行目が見出し、2行目以降がデータで、A〜I列を読み込みます
- 必要な列: `預り日`、`患者ID`、`文書名`、`診療科`、`医師名`、`医師依頼日`、`担当者名`
- 元データとデータベースの列(見出し)は同じである必要があります

## データ処理の流れ
1. データ読込
   - 元データとデータベースを結合します
   - 医師依頼日または担当者名が空の行を除きます
   - 預り日・患者ID・文書名・診療科・医師名が同じ行を重複として除きます
   - 預り日・診療科・患者IDの順に並べてデータベースに保存し、`backup_dir` に `backup_{ファイル名}` としてコピーします
2. 集計開始
   - データベースから、預り日が期間内(両端を含む)の行を取り出します
   - 担当者×診療科、診療科ごと、全体の件数を数えてテンプレートに書き込みます

## 設定ファイル
設定は `utils/config.ini` の1つだけです。ビルドした exe では、実行ファイルに同梱された `config.ini` を読み書きします。

| セクション | キー | 内容 |
|---|---|---|
| `[Appearance]` | `window_width` / `window_height` | ウィンドウの幅・高さ(px) |
| `[PATHS]` | `source_file_path` | 元データ(医療文書担当一覧)のパス |
| | `database_path` | 医療文書担当一覧データベースのパス |
| | `backup_dir` | データベースのバックアップ先フォルダ |
| | `template_path` | 出力書式(テンプレートExcel)のパス |
| | `output_dir` | 集計結果の出力先フォルダ |
| `[Analysis]` | `start_date` / `end_date` | 前回の分析期間(アプリが自動で更新) |
| | `year_start_month` | 年の始まりの月(1 または 4)(アプリが自動で更新) |
| | `ordered_names` | 集計表に並べる担当者(カンマ区切り) |
| | `clinical_departments` | 集計表に並べる診療科(カンマ区切り。`合計` を入れるとその列に行の合計を出力) |

## 集計結果
- A1: `医療文書作成件数 {最初の預り日}-{最後の預り日}`
- 2行目: `氏名` と `clinical_departments` の診療科
- 3行目以降: `ordered_names` の担当者ごとの件数
- 最終行: 診療科ごとの合計(一覧にない担当者の件数も含む)。`合計` 列は全体の件数

出力ファイル名は `医療文書作成件数{最初の預り日}-{最後の預り日}.xlsx` の形式です(例: `医療文書作成件数20250401-20250630.xlsx`)。

## プロジェクト構成
```
MediDocsAnalyzer/
├── main.py                               # エントリーポイント
├── app/
│   ├── __init__.py                       # バージョン情報(__version__)
│   ├── app_window.py                     # メイン画面(MedicalDocsAnalyzerGUI)
│   ├── path_settings_dialog.py           # ファイルパス設定ダイアログ
│   └── constants.py                      # 画面に表示するメッセージ
├── service/
│   ├── service_medical_docs_analyzer.py  # 読込・集計全体の統括
│   ├── service_excel_reader.py           # Excelの読み込み(全セルを文字列に統一)
│   ├── service_data_analyzer.py          # Polars による結合・重複排除・集計
│   └── service_excel_writer.py           # データベース・バックアップ・集計結果の書き出し
├── utils/
│   ├── config.ini                        # 設定ファイル
│   ├── config_manager.py                 # 設定ファイルの読み書き
│   └── date_period.py                    # 年・期間から開始日・終了日を求める
├── assets/                               # アプリのアイコン
├── tests/                                # テスト(ソースと同じ構成)
└── docs/                                 # CHANGELOG.md・LICENSE
```

## 開発者向け情報
### 開発コマンド
```bash
uv run pyright                       # 型チェック
uv run pytest tests/ -v --tb=short  # テスト(全件)
uv run python build.py               # exe のビルド
```

### コードの構成
- `MedicalDocsAnalyzerGUI`: GUIの作成と管理を担当
- `PathSettingsDialog`: ファイルパスの設定ダイアログ
- `MedicalDocsAnalyzer`: データ読込(`load_data`)と集計(`run_analysis`)の統括。エラーは GUI へ送出
- `service_excel_reader`: Excelファイルの読み込み(`read_documents`)と `NoDataError`
- `service_data_analyzer`: 結合・重複排除(`merge_documents`)、期間の絞り込み(`filter_by_period`)、件数の集計(`count_documents`)
- `service_excel_writer`: データベースの書き換え(`write_database`)、バックアップ(`backup_file`)、集計結果の出力(`save_results`)

### データフロー
1. `MedicalDocsAnalyzerGUI` がユーザー入力を受け取り、`MedicalDocsAnalyzer.load_data` / `run_analysis` を呼び出す
2. `MedicalDocsAnalyzer` が設定を読み込み、`read_documents` でデータを読み込む
3. `service_data_analyzer` が結合・絞り込み・集計を行う
4. `service_excel_writer` がデータベースまたは集計結果を書き出す

## トラブルシューティング
1. 「元データにデータがありません。」と表示される
   - 「ファイルパス」で元データのパスが正しいか確認
2. 「元データとデータベースの列が一致しません」と表示される
   - 元データとデータベースの1行目(見出し)が同じか確認
3. 「指定された期間内のデータがありません。」と表示される
   - 先に「データ読込」を実行したか確認
   - 預り日が `yyyy/mm/dd` 形式になっているか確認
4. ファイルが保存できない
   - データベースや同名の結果ファイルをExcelで開いたままだと保存できないため、閉じてから再実行

## ライセンス
このプロジェクトのライセンス情報については、[LICENSE](docs/LICENSE) を参照してください。

## 更新履歴
更新履歴は [CHANGELOG.md](docs/CHANGELOG.md) を参照してください。
