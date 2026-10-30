# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Windows専用のtkinter GUIアプリ。医療文書の記録をExcelから読み込み、担当者・診療科ごとに集計してExcelテンプレートに出力する。外部API・LLM・環境変数は使わない。

## 構成

- `app/`: GUI(`MedicalDocsAnalyzerGUI`)。右側の `ResultView` が「全体」「個人」タブで表と matplotlib の図(`charts.py`)を出す。ほかに `MasterDialog`、`ImportErrorsDialog`、`PathSettingsDialog`、`constants.py`(UI文言はすべてここに定数で置く)。
- 印刷は Excel 出力(`service_excel_writer`、openpyxl のグラフ付き・A4 横 1 枚)を Excel で開いて行う。tkinter に印刷機能がないため。
- `service/service_medical_docs_analyzer.py`: `MedicalDocsAnalyzer` が統括する。データ読込は `prepare_import()` で計画(`ImportPlan`)を作り、GUI で確認してから `apply_import()` で書く。例外は握りつぶさず GUI まで送る(`NoDataError` は専用メッセージを出す)。
- `service/`: `service_excel_reader`(見出し名で列を探し、`Document` とエラー行 `RowError` に分ける)、`service_database`(SQLite の接続・スキーマ・文書の入れ替え)、`service_data_analyzer`(SQL で集計)、`service_backup`(起動時の日次バックアップ)、`service_excel_writer`(テンプレート出力)。
- `utils/`: `config_manager.py`、`date_period.py`(年・期間・年の始まりの月から開始日と終了日を求める)。
- DB は SQLite(共有フォルダに置くため `journal_mode=DELETE`)。日付は ISO 形式の文字列。患者名・患者ID・備考は取り込まない。
- 取り込みは範囲置き換え: 元データの預り日(下書きの行も含む)の最も古い日〜新しい日にある DB の文書を消して入れ直す。元データが古いコピーでも、それより新しい DB の文書は残る。元データで削除した行(取り消し)も DB から消える。DB が空なら先に `legacy_database_path`(旧 xlsx の DB)を取り込む。
- 文書にするのは医師依頼日と担当者名がある行。空なら下書きとして飛ばす。日付が読めない・医師依頼日が預り日より前の行はエラー行として除き、一覧に出す。
- 集計対象は預り日が期間内の文書。
- 担当者・診療科はDBのマスタ(`service_master`)。`config.ini` の `ordered_names` / `clinical_departments` はマスタが空のときの初期値にだけ使う。取り込み時にマスタにない担当者名はエラー行にする。

## 設定

- 設定はすべて `utils/config.ini`(utf-8)。`[PATHS]` は `C:\Shinseikai\MediDocsAnalyzer\...` の絶対パス固定。
- 凍結(PyInstaller)時は exe と同じフォルダの ini を使う。無ければ `sys._MEIPASS` に同梱した ini をコピーする。
- テストは `tests/conftest.py` の `config_path` フィクスチャで一時 ini に差し替える。本物の `config.ini` を書き換えるテストを書かない。

## 注意点

- バージョンは `app/__init__.py` の `__version__`、`pyproject.toml` の `version`、`docs/CHANGELOG.md` の3か所。
- `build.py` は PyInstaller 呼び出しのスタブ(アプリ名はプレースホルダ、`--add-data` の区切りは Windows では `;`、アイコン未指定)。
- UI文言・コメント・データはすべて日本語。ファイル入出力は utf-8 を明示する(`ordered_names` に旧字体「﨑」を含む)。
