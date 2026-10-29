# 変更履歴

このプロジェクトのすべての重要な変更は、このファイルに記録されます。

フォーマットは [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に基づいており、
バージョン番号は [Semantic Versioning](https://semver.org/lang/ja/) に従っています。

## [Unreleased]

## [1.1.0] - 2026-10-29

### 追加
- 年・年の始まり(1月始まり/4月始まり、初期値は4月始まり)・期間(月・四半期・上半期・下半期・一年)を選んで分析期間を自動入力する機能
- 元データ・データベース・バックアップ先・出力書式・出力先を画面から設定する「ファイルパス」ダイアログ
- データ読込の完了時に、データベースの件数を表示

### 変更
- 設計を作り直し、`service` を読み込み(`service_excel_reader`)・集計(`service_data_analyzer`)・書き出し(`service_excel_writer`)・統括(`service_medical_docs_analyzer`)に分割
- 処理中のエラーをコンソールに出して握りつぶすのをやめ、画面にエラーメッセージを表示するように変更
- 期間内にデータがない場合は、空の集計表を出力せずにメッセージを表示するように変更
- 元データとデータベースの列が一致しない場合は、データベースを元データだけで上書きせず、エラーを表示するように変更
- 日付を文字列ではなく日付型で扱い、`yyyy/mm/dd` 以外の手入力はエラーとして表示するように変更
- 「設定」ボタンは実行中の `config.ini` を開くように変更し、`[PATHS]` の `config_path` を廃止
- 画面の文言を `app/constants.py` にまとめ、全関数に型ヒントを追加

### 削除
- 未使用の `utils/utils.py`
- `service_data_processor.py`、`service_excel_handler.py`、`service_medical_docs_processor.py`(新しいモジュールへ統合)

### 修正
- `app/__init__.py` と `pyproject.toml` のバージョンの不一致

## [1.0.0] - 2026-10-28
- 初版リリース
