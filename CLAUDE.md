# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Windows専用のtkinter GUIアプリ。医療文書の記録をExcelから読み込み、担当者・診療科ごとに集計してExcelテンプレートに出力する。外部API・LLM・環境変数は使わない。

## 構成

- `app/`: GUI(`MedicalDocsAnalyzerGUI`)、`PathSettingsDialog`、`constants.py`(UI文言はすべてここに定数で置く)。
- `service/service_medical_docs_analyzer.py`: `MedicalDocsAnalyzer` が統括する。`load_data()` は元データをDBに取り込んでからバックアップし、`run_analysis(start, end)` は集計結果を出力してパスを返す。例外は握りつぶさず GUI まで送る(`NoDataError` は専用メッセージを出す)。
- `service/`: `service_excel_reader`(全セルを文字列に統一し、日付は `YYYY/MM/DD`、患者IDは数値表記にそろえる)、`service_data_analyzer`(polars で結合・重複排除・期間フィルタ・集計)、`service_excel_writer`(DB書き換え・バックアップ・テンプレート出力)。
- `utils/`: `config_manager.py`、`date_period.py`(年・期間・年の始まりの月から開始日と終了日を求める)。
- 集計対象の条件: 医師依頼日と担当者名が空でなく、預り日が期間内の行。預り日・患者ID・文書名・診療科・医師名で重複排除する。
- DB ファイルの形式(A〜I列、日付は `YYYY/MM/DD` 文字列、患者IDは数値)は既存ファイルとの互換のため変えない。

## 設定

- 設定はすべて `utils/config.ini`(utf-8)。`[PATHS]` は `C:\Shinseikai\MediDocsAnalyzer\...` の絶対パス固定。
- 凍結(PyInstaller)時は `sys._MEIPASS` 配下の ini を使うため、`save_config` の書き込みは永続化されない。
- テストは `tests/conftest.py` の `config_path` フィクスチャで一時 ini に差し替える。本物の `config.ini` を書き換えるテストを書かない。

## 注意点

- バージョンは `app/__init__.py` の `__version__`、`pyproject.toml` の `version`、`docs/CHANGELOG.md` の3か所。
- `build.py` は PyInstaller 呼び出しのスタブ(アプリ名はプレースホルダ、`--add-data` の区切りは Windows では `;`、アイコン未指定)。
- UI文言・コメント・データはすべて日本語。ファイル入出力は utf-8 を明示する(`ordered_names` に旧字体「﨑」を含む)。
