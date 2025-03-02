# 業務分析アプリ

## 概要
業務分析アプリは、日々の業務内容と時間を記録したWILLDOリストを分析し、業務内容の集計を行うツールです。
クラーク業務、デイリータスク、コミュニケーションなど、様々な観点から業務時間の分析が可能です。

## 動作要件
- Python 3.11以上
- Windows OS
- 必要なPythonパッケージ:
  - tkcalendar
  - polars
  - openpyxl
  - pyarrow

## インストール方法
1. リポジトリをクローンまたはダウンロードします
2. 必要なパッケージをインストールします：
```bash
pip install -r requirements.txt
```

## アプリケーションの構成
- `main.py`: アプリケーションのエントリーポイント
- `app_window.py`: GUIの実装
- `service_task_analyzer.py`: 分析の全体的な処理の実装
- `service_excel_reader.py`: Excelファイルの読み込み処理
- `service_data_analyzer.py`: データの集計・分析ロジック
- `service_excel_writer.py`: 分析結果のExcel出力処理
- `utils.py`: ユーティリティ関数
- `config_manager.py`: 設定ファイル管理
- `version.py`: アプリケーションのバージョン情報

## 使用方法
アプリケーションの起動：
```bash
python main.py
```

1. GUIで分析期間（開始日・終了日）を選択します
2. 「分析開始」ボタンをクリックすると分析が実行されます
3. 分析結果は指定された出力フォルダにExcelファイルとして保存されます
4. 「設定ファイル」ボタンをクリックすると、設定ファイルが開きます

## 設定ファイル
設定ファイルには以下の項目を定義する必要があります：

### [PATHS]セクション
- `input_file_path`: WILLDOリストのExcelファイルパス
- `template_path`: 出力テンプレートのパス
- `output_dir`: 分析結果の出力先ディレクトリ
- `config_path`: 設定ファイルのパス

### [Analysis]セクション
- `start_row`: 業務データの開始行
- `end_row`: 業務データの終了行
- `daily_task_start_row`: デイリータスクの開始行
- `daily_task_end_row`: デイリータスクの終了行
- `communication_start_row`: コミュニケーションデータの開始行
- `communication_end_row`: コミュニケーションデータの終了行

### [Appearance]セクション
- `window_width`: ウィンドウの幅
- `window_height`: ウィンドウの高さ

## 分析結果
分析結果は以下の項目を含むExcelファイルとして出力されます：

1. クラーク業務の集計
2. クラーク以外の業務の集計
3. デイリータスクの集計
4. コミュニケーションの集計（氏名と時間）
5. コミュニケーション内容の集計
6. 全項目の集計

出力ファイル名は「WILLDOリストまとめ{開始日}_{終了日}.xlsx」の形式になります。
分析完了後、自動的にExcelで結果ファイルが開かれます。

## 開発者向け情報
### コードの構成
- `TaskAnalyzerGUI`: GUIの作成と管理を担当
- `TaskAnalyzer`: 分析処理全体の統括
- `ExcelTaskReader`: Excelファイルからのデータ読み込み
- `TaskDataAnalyzer`: Polarsを使ったデータフレーム操作と集計
- `ExcelResultWriter`: 分析結果のExcel出力

### データフロー
1. `TaskAnalyzerGUI`がユーザー入力を受け取り`TaskAnalyzer`に処理を依頼
2. `TaskAnalyzer`は`ExcelTaskReader`を使ってデータを読み込み
3. 読み込んだデータを`TaskDataAnalyzer`を使って分析
4. 分析結果を`ExcelResultWriter`を使ってExcelファイルに出力

### 拡張方法
1. 新しい分析項目の追加
   - `service_data_analyzer.py`の`analyze_task_data`メソッドに新しい分析ロジックを追加
   - テンプレートExcelファイルに新しいシートを追加
   - `service_excel_writer.py`の`save_results`メソッドに新しいシートへの書き込み処理を追加

2. GUIの拡張
   - `app_window.py`の`_setup_gui`メソッドを修正
   - 必要に応じて新しいフレームやコントロールを追加

3. データ処理の拡張
   - `service_excel_reader.py`を拡張して新しいデータの読み込み処理を追加
   - `utils.py`に共通のユーティリティ関数を追加

## 実装の詳細
### データ処理
- Polarを使ったデータフレーム操作による高速な集計処理
- 正規表現を使った文字列からの名前抽出機能
- エラーハンドリングと安全な型変換

### GUI
- tkcalendarを使用した日付選択UI
- 設定ファイルからのウィンドウサイズ読み込み
- エラーメッセージのポップアップ表示

## トラブルシューティング
1. アプリケーションが起動しない
   - Python 3.11以上がインストールされているか確認
   - 必要なパッケージがすべてインストールされているか確認
   - `pip install -r requirements.txt`を実行して必要なパッケージをインストール

2. 分析が実行されない
   - 設定ファイルの内容が正しいか確認
   - 入力ファイルが指定されたパスに存在するか確認
   - 入力ファイルの形式が正しいか確認

3. エラーメッセージが表示される
   - エラーメッセージの内容を確認
   - 設定ファイルの値が正しいか確認
   - 入力ファイルの内容が正しい形式か確認

4. 分析結果が空または不完全
   - 指定された期間内にデータが存在するか確認
   - 設定ファイルの行指定が正しいか確認

## ライセンス
このプロジェクトのライセンス情報については、LICENSEファイルを参照してください。
