# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Windows専用のtkinter GUIアプリ。別のアプリ CSV2XL が作る医療文書の SQLite(`C:\Shinseikai\CSV2XL\医療文書担当一覧.db`)を読み取り専用で読み、担当者・診療科ごとに集計して Excel に出力する。文書の取り込みと DB のバックアップは CSV2XL が行い、このアプリでは行わない。外部API・LLM・環境変数は使わない。

## 構成

- `app/`: GUI(`MedicalDocsAnalyzerGUI`)。右側の `ResultView` が「概要」タブで下書きの一覧を、「全体」「個人」「全体推移」「個人推移」タブで表と matplotlib の図(`charts.py`)を出す。推移の 2 タブは `TrendPanel` を共用する。ほかに `MasterDialog`、`SettingsDialog`(ファイルパス・画面の大きさ・概要で確認する日数・個人の Excel 出力の文字の大きさ。`[Analysis]` はメイン画面とマスタと重なるので出さない)、`date_entry.py`(日付入力欄と開始日・終了日の読み取り)、`constants.py`(UI文言はすべてここに定数で置く)。`MedicalDocsAnalyzerGUI` は `__init__` で画面を作るだけで、起動時の処理(概要・推移・前回の期間の集計)は `main.py` が呼ぶ `start()` で行う。
- 印刷は Excel 出力(`service_excel_writer`)を Excel で開いて行う。tkinter に印刷機能がないため。Excel 出力は `output_dir` に保存し、ファイルではなくフォルダを開く(GUI 側の `app/folder.py` の `open_folder_of`。writer は保存してパスを返すだけ。エクスプローラーで同じフォルダが開いていれば、その窓を使う)。
- 設定の保存は `utils/config_manager.update_config` を使う(ほかの画面の変更を消さないよう、保存直前に読み直す)。
- `service/service_medical_docs_analyzer.py`: `MedicalDocsAnalyzer` が統括する。文書は `open_documents_database()`(`[PATHS] documents_database_path`)、担当者・診療科・ログインは `_open_database()`(`[PATHS] database_path`)で開く。「集計開始」は取り込みをせず、文書の DB の今の内容で集計し、概要も読み直す。例外は握りつぶさず GUI まで送る(`NoDataError` は専用メッセージを出す)。
- `service/`: `models`(`Draft`・`NoDataError` などの共通のデータ型)、`service_database`(マスタの DB の接続・スキーマと、文書の DB の読み取り専用の接続 `open_documents`。預り日の期間の条件 `PERIOD_CONDITION`)、`service_data_analyzer`(SQL で集計。全体の表の行は `OverallSummary.rows` で画面と Excel が共用する)、`service_overview`(概要の下書きの読み込みと状態・件数)、`service_excel_writer`(Excel 出力)。
- `utils/`: `config_manager.py`、`date_period.py`(年・期間・年の始まりの月から開始日と終了日を求める)、`log_rotation.py`(ログは `config.ini` と同じフォルダの `[LOGGING] log_directory`)、`file_cleaner.py`(`main.py` が起動時に `output_dir` の `[FileCleanup] retention_days` 日を過ぎたファイルを消す。失敗はログだけ。output は端末ごと)。
- 文書の DB(CSV2XL): `file:...?mode=ro` の URI で開き、書き込まない。無いファイルは作らずにエラーにする。スキーマは CSV2XL の `services/schema.sql`(版 2。`patient_id` は整数、`staff_name` と `requested_date` は NULL あり、`deleted` は論理削除)。集計の SQL は TEMP VIEW `valid_documents`(`deleted = 0` で担当者名と医師依頼日の両方がある行)から読む。条件は `VALID_DOCUMENTS_VIEW` 1 か所。医師依頼日が預り日より前の行や担当者名のリストにない名前は CSV2XL が取り込み時にはじくので、このアプリでは確かめない。
- マスタの DB: SQLite(共有フォルダに置くため `journal_mode=DELETE`)。担当者・診療科・ログイン情報だけを持つ。スキーマの版は `PRAGMA user_version`(`SCHEMA_VERSION`)で持ち、古い版は `migrate` でそろえ(版 2 で取り込んでいた documents テーブルを DROP)、新しい版(C# 移植版など)は開かない。
- 集計対象は預り日が期間内の文書。
- 概要: 文書の DB から、預り日が本日を含め `[Overview] pending_days` 日以内で、`deleted = 0` かつ担当者名か医師依頼日が空の行を下書き(`Draft`)として読む(`service_overview.load_overview`、マスタの DB は開かない)。起動時・集計開始・設定の保存後に読み直す。状態は `draft_status`(両方空=未着手、医師依頼日だけ空=未完了、担当者名だけ空=担当不明)で、診療科ごとの件数表と行の一覧を出す。読めなくてもダイアログは出さずタブ内に出す。職員にも見せる。
- 起動時に `[Analysis] start_date / end_date` があれば `analyze_saved_period` で集計開始と同じ処理を自動で行う。
- 集計対象(`DocumentType`): 全書類・医療文書(主治医意見書を除く)・主治医意見書。主治医意見書は文書名に「主治医意見書」を含む文書で、SQL の条件は `document_type_condition` 1 か所。左側の選択欄を選び直すと `change_document_type` が `MedicalDocsAnalyzerGUI.document_type` に確定して `[Analysis] document_type` に保存し(無いか不正なら全書類)、推移と、集計済み(`period` がある)なら全体・個人を DB から集計し直す。全体・個人・推移・Excel 出力(表題とファイル名に名前を出す)すべてに効かせる。推移の表題は `trend_subject`(「[氏名] 集計対象 [診療科]」)。「合計」「主治医意見書」は診療科の名前に登録させない。
- `tests/fixtures/source_golden.json` は文書の DB の行と、その集計・概要の見本(C# 移植版と結果を比べるため)。集計や概要を変えたら更新する。テスト用の文書の DB は `tests/documents_database.py` で作る(スキーマは CSV2XL の写し)。
- Excel 出力のテンプレート: 全体は `医療文書作成件数.xlsx`、全体推移・個人推移は `医療文書作成件数推移.xlsx`(`[PATHS] template_dir`。テスト用の写しは `tests/fixtures/templates/`)。先頭のシートの表に件数だけを入れ、グラフはテンプレートのグラフシート(Excel で作ったもの。openpyxl で読み書きしても残る)が描く。2 行目が見出し、3 行目から A 列の名前(氏名・年)で行を探し、テンプレートに無い担当者・診療科・年は出さない。A1 の `{集計対象}` などを置き換えた表題をファイル名にも使う(全体推移は `_{氏名}` ごと消す)。推移は月の見出しと年の名前を年の始まりの月に合わせて書き換え、0 件の月は空欄にする。テンプレートが無ければ `FileNotFoundError` で、GUI が専用のメッセージを出す。
- 個人タブは表だけ(グラフなし)で、文書名は全件(スクロール)。個人の Excel 出力はテンプレートを使わずコードで作り(グラフなし)、文書名は上位 `TOP_DOCUMENT_COUNT` 件、表の文字の大きさは `[Excel] staff_font_size`(既定 12、設定画面で変える)。
- 全体推移(`count_trend`): 医師依頼日が範囲(開始日〜終了日。左側の分析期間とは連動しない)内の DB の文書を、年(年の始まりの月は左側の「1月/4月始まり」)ごとの月別に数え、1 年 1 本の折れ線にする。範囲外と本日より後の月は None(空欄)、今月は途中の件数。0 件の月は線を引かない(Excel は表の 0 を空欄にする)。前後が 0 件の月も見えるよう点の印を付ける。画面のグラフの件数の軸は整数の目盛りだけにする。診療科は選択欄(合計・マスタの診療科)で選ぶ。選択欄の文言と集計対象は `TrendRequest.document_filter` で `DocumentFilter` にしてサービスへ渡す(文字列を条件に使わない)。開始日・終了日・診療科は `[Trend]` に保存し、無ければ前年(年度)の初日〜本日の合計。職員に見せるかは `[Login] staff_tabs`(判定は `MedicalDocsAnalyzerGUI.shows_trend`)。
- 個人推移: 全体推移を担当者で絞ったもの(`count_trend` の `staff_name`)。氏名は退職者も含むマスタの全担当者。選択は `[StaffTrend]` に保存する。職員は本人だけ見られる(判定は `MedicalDocsAnalyzerGUI.restricted_staff` 1 か所)。
- ログイン(`service_auth`、`app/login_dialog.py`): `[Login] enabled`(端末ごと、既定 true)が false ならログインなしですべてを見せる。true で、パスワードを設定した在籍中の管理者が 1 人でもいれば、起動時に電子カルテIDとパスワードを求める。パスワードは PBKDF2 のハッシュで保存する。仮パスワードはマスタ画面で発行し、最初のログインで変更させる。管理者権限はマスタの担当者の `role`。職員ロールは `[Login] staff_tabs`(タブの表示名のカンマ区切り、既定は概要・全体推移・個人・個人推移。`DEFAULT_STAFF_TABS`)のタブだけを見られ、個人・個人推移は本人だけ。マスタ・設定は使えない(集計開始はできる)。`[Login]` は設定画面に出さない。マスタの DB を開けないときはファイルパスを直せるようログインなしで起動する。
- 担当者・診療科はマスタの DB(`service_master`)。`config.ini` の `ordered_names` / `clinical_departments` はマスタが空のときの初期値にだけ使う。担当者名は CSV2XL の担当者名のリスト(`list_items` の staff)とは別に管理するので、両方に登録が要る。全体の表に並ぶのはマスタの担当者・診療科だけ。

## 設定

- 設定はすべて `utils/config.ini`(utf-8)。`[PATHS]` は `C:\Shinseikai\...` の絶対パス固定(`documents_database_path` は CSV2XL の DB、`database_path` はマスタの DB、`template_dir` は Excel 出力のテンプレートのフォルダ)。
- 凍結(PyInstaller)時は exe と同じフォルダの ini を使う。無ければ `sys._MEIPASS` に同梱した ini をコピーする。
- テストは `tests/conftest.py` の `config_path` フィクスチャで一時 ini に差し替える。本物の `config.ini` を書き換えるテストを書かない。

## 注意点

- バージョンは `app/__init__.py` の `__version__`、`pyproject.toml` の `version`、`docs/CHANGELOG.md` の3か所。
- `build.py` は PyInstaller 呼び出しのスタブ(アプリ名はプレースホルダ、`--add-data` の区切りは Windows では `;`、アイコン未指定)。
- UI文言・コメント・データはすべて日本語。ファイル入出力は utf-8 を明示する(`ordered_names` に旧字体「﨑」を含む)。
