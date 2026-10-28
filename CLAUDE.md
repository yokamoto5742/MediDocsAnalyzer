# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Windows専用のtkinter GUIアプリ。医療文書の記録をExcelから読み込み、担当者・診療科ごとに集計してExcelテンプレートに出力する。外部API・LLM・環境変数は使わない。

## 構成

- `app/app_window.py`: GUI(`MedicalDocsAnalyzerGUI`)。処理の流れは GUI → `load_config` → `service_medical_docs_processor`(元データをDBへ取り込み、事前にバックアップ) → `service_medical_docs_analyzer`(集計・テンプレート出力)。
- `service/`: 業務ロジック。`service_data_processor`(セル・日付変換、期間フィルタ)、`service_excel_handler`(Excel読み書き・バックアップ)。
- `utils/`: `config_manager.py` と `utils.py`。
- 集計対象の条件: 医師依頼日と担当者名が空でなく、預り日が期間内の行。預り日・患者ID・文書名・診療科・医師で重複排除する。

## 設定

- 設定はすべて `utils/config.ini`(utf-8)。`[PATHS]` は `C:\Shinseikai\MediDocsAnalyzer\...` の絶対パス固定。
- 凍結(PyInstaller)時は `sys._MEIPASS` 配下の ini を使うため、`save_config` の書き込みは永続化されない。

## 注意点

- バージョンは `app/__init__.py` の `__version__`、`pyproject.toml` の `version`、`docs/CHANGELOG.md` の3か所。
- `build.py` は PyInstaller 呼び出しのスタブ(アプリ名はプレースホルダ、`--add-data` の区切りは Windows では `;`、アイコン未指定)。
- UI文言・コメント・データはすべて日本語。ファイル入出力は utf-8 を明示する(`ordered_names` に旧字体「﨑」を含む)。
