# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Windows専用のtkinter GUIアプリ。医療文書の記録をExcelから読み込み、担当者・診療科ごとに集計してExcelテンプレートに出力する。外部API・LLM・環境変数は使わない。

## 構成

- `app/`: GUI(`MedicalDocsAnalyzerGUI`)。右側の `ResultView` が「概要」タブで下書きの一覧を、「全体」「個人」「全体推移」「個人推移」タブで表と matplotlib の図(`charts.py`)を出す。推移の 2 タブは `TrendPanel` を共用する。ほかに `MasterDialog`、`ImportErrorsDialog`、`SettingsDialog`(ファイルパス・画面の大きさ・バックアップ世代数・概要で確認する日数。`[Analysis]` はメイン画面とマスタと重なるので出さない)、`date_entry.py`(日付入力欄と開始日・終了日の読み取り)、`constants.py`(UI文言はすべてここに定数で置く)。`MedicalDocsAnalyzerGUI` は `__init__` で画面を作るだけで、起動時の処理(バックアップ・概要・推移・前回の期間の集計)は `main.py` が呼ぶ `start()` で行う。
- 印刷は Excel 出力(`service_excel_writer`、openpyxl のグラフ付き・A4 横 1 枚)を Excel で開いて行う。tkinter に印刷機能がないため。Excel 出力と取込エラー行の CSV 保存は `output_dir` に保存し、ファイルではなくフォルダを開く(GUI 側の `app/folder.py` の `open_folder_of`。writer は保存してパスを返すだけ。エクスプローラーで同じフォルダが開いていれば、その窓を使う)。
- 設定の保存は `utils/config_manager.update_config` を使う(ほかの画面の変更を消さないよう、保存直前に読み直す)。
- `service/service_medical_docs_analyzer.py`: `MedicalDocsAnalyzer` が統括する。取り込みは `prepare_source_import()`(元データ xlsm。`[PATHS] source_dir` のフォルダで、ファイル名の末尾 `_YYYYMMDDhhmm.xlsm` の日時が最新のもの。`find_latest_source`)/ `prepare_legacy_import()`(旧 DB xlsx)で計画(`ImportPlan`)を作り、`apply_import()` で書く。「集計開始」は元データを確認なしで取り込んでから集計し、「Excel読込」は旧 DB を確認してから取り込む。例外は握りつぶさず GUI まで送る(`NoDataError` は専用メッセージを出す)。
- `service/`: `models`(`Document`・`Draft`・`RowError`・`SourceData`・`NoDataError` などの共通のデータ型)、`service_excel_reader`(見出し名で列を探し、`Document` とエラー行 `RowError` に分ける)、`service_database`(SQLite の接続・スキーマ・文書の入れ替え。預り日の期間の条件 `PERIOD_CONDITION`)、`service_data_analyzer`(SQL で集計。全体の表の行は `OverallSummary.rows` で画面と Excel が共用する)、`service_overview`(概要の下書きの状態と件数)、`service_backup`(起動時の日次バックアップ)、`service_excel_writer`(テンプレート出力)。
- `utils/`: `config_manager.py`、`date_period.py`(年・期間・年の始まりの月から開始日と終了日を求める)、`log_rotation.py`(ログは `config.ini` と同じフォルダの `[LOGGING] log_directory`)、`file_cleaner.py`(`main.py` が起動時に `output_dir` の `[FileCleanup] retention_days` 日を過ぎたファイルを消す。失敗はログだけ。output は端末ごと)。
- DB は SQLite(共有フォルダに置くため `journal_mode=DELETE`)。日付は ISO 形式の文字列。院内だけで使うため患者IDは取り込む。スキーマの版は `PRAGMA user_version`(`SCHEMA_VERSION`)で持ち、古い版は `migrate` でそろえ、新しい版(C# 移植版など)は開かない。患者名・備考は取り込まない。
- 取り込みは範囲置き換え: 読んだファイルの預り日(下書きの行も含む)の最も古い日〜新しい日にある DB の文書を消して入れ直す。ファイルが古いコピーでも、それより新しい DB の文書は残る。ファイルで削除した行(取り消し)も DB から消える。
- 文書にするのは医師依頼日と担当者名がある行。空なら下書きとして飛ばす。日付が読めない・医師依頼日が預り日より前の行はエラー行として除き、一覧に出す。
- 集計対象は預り日が期間内の文書。
- 概要: 下書きは DB に入れず、`SourceData.drafts` で返す。起動時(`load_overview`、DB は開かない)・集計開始(取り込んだ計画の drafts)・設定の保存後に、預り日が本日を含め `[Overview] pending_days` 日以内の下書きを出す。状態は `draft_status`(両方空=未着手、医師依頼日だけ空=未完了、担当者名だけ空=担当不明)で、診療科ごとの件数表と行の一覧を出す。起動時に読めなくてもダイアログは出さずタブ内に出す。職員にも見せる。
- 起動時に `[Analysis] start_date / end_date` があれば `analyze_saved_period` で集計開始と同じ処理を自動で行う。煩わしいため、起動時は取り込めなかった行の一覧(`ImportErrorsDialog`)を出さない(`show_errors=False`)。
- 集計対象(`DocumentType`): 全書類・医療文書(主治医意見書を除く)・主治医意見書。主治医意見書は文書名に「主治医意見書」を含む文書で、SQL の条件は `document_type_condition` 1 か所。左側の選択欄は「集計開始」で `MedicalDocsAnalyzerGUI.document_type` に確定して `[Analysis] document_type` に保存し(無いか不正なら全書類)、全体・個人・推移・Excel 出力(表題とファイル名に名前を出す)すべてに効かせる。推移の表題は `trend_subject`(「[氏名] 集計対象 [診療科]」)。「合計」「主治医意見書」は診療科の名前に登録させない。
- `tests/fixtures/source_golden.json` は元データの読み取りと集計の見本(C# 移植版と結果を比べるため)。読み取りや集計を変えたら更新する。
- 個人タブの文書名は全件(スクロール)、個人の Excel 出力は上位 `TOP_DOCUMENT_COUNT` 件。全体の Excel 出力は表だけで、グラフは出さない。
- 全体推移(`count_trend`): 元データは取り込まず、医師依頼日が範囲(開始日〜終了日。左側の分析期間とは連動しない)内の DB の文書を、年(年の始まりの月は左側の「1月/4月始まり」)ごとの月別に数え、1 年 1 本の折れ線にする。範囲外と本日より後の月は None(空欄)、今月は途中の件数。0 件の月は線を引かない(Excel は隠しシート「グラフ用」で 0 を空欄にする)。前後が 0 件の月も見えるよう点の印を付ける。診療科は選択欄(合計・マスタの診療科)で選ぶ。選択欄の文言と集計対象は `TrendRequest.document_filter` で `DocumentFilter` にしてサービスへ渡す(文字列を条件に使わない)。開始日・終了日・診療科は `[Trend]` に保存し、無ければ前年(年度)の初日〜本日の合計。職員には見せない。Excel のグラフは凡例・タイトルの `overlay=False` を明示する(無いと Excel が重ねて表示する)。
- 個人推移: 全体推移を担当者で絞ったもの(`count_trend` の `staff_name`)。氏名は退職者も含むマスタの全担当者。選択は `[StaffTrend]` に保存する。職員は本人だけ見られる(判定は `MedicalDocsAnalyzerGUI.restricted_staff` 1 か所)。
- ログイン(`service_auth`、`app/login_dialog.py`): パスワードを設定した在籍中の管理者が 1 人でもいれば、起動時に電子カルテIDとパスワードを求める。パスワードは PBKDF2 のハッシュで保存する。仮パスワードはマスタ画面で発行し、最初のログインで変更させる。職員ロールは自分の個人集計だけを見られ、Excel読込・マスタ・設定は使えない(集計開始での元データの取り込みはできる)。DB を開けないときはファイルパスを直せるようログインなしで起動する。
- 担当者・診療科はDBのマスタ(`service_master`)。`config.ini` の `ordered_names` / `clinical_departments` はマスタが空のときの初期値にだけ使う。取り込み時にマスタにない担当者名はエラー行にする。

## 設定

- 設定はすべて `utils/config.ini`(utf-8)。`[PATHS]` は `C:\Shinseikai\MediDocsAnalyzer\...` の絶対パス固定。
- 凍結(PyInstaller)時は exe と同じフォルダの ini を使う。無ければ `sys._MEIPASS` に同梱した ini をコピーする。
- テストは `tests/conftest.py` の `config_path` フィクスチャで一時 ini に差し替える。本物の `config.ini` を書き換えるテストを書かない。

## 注意点

- バージョンは `app/__init__.py` の `__version__`、`pyproject.toml` の `version`、`docs/CHANGELOG.md` の3か所。
- `build.py` は PyInstaller 呼び出しのスタブ(アプリ名はプレースホルダ、`--add-data` の区切りは Windows では `;`、アイコン未指定)。
- UI文言・コメント・データはすべて日本語。ファイル入出力は utf-8 を明示する(`ordered_names` に旧字体「﨑」を含む)。
