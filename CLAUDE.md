# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Windows専用のtkinter GUIアプリ。医療文書の記録をExcelから読み込み、担当者・診療科ごとに集計してExcelテンプレートに出力する。外部API・LLM・環境変数は使わない。

## 構成

- `app/`: GUI(`MedicalDocsAnalyzerGUI`)。右側の `ResultView` が「概要」タブで下書きの一覧を、「全体」「個人」「全体推移」タブで表と matplotlib の図(`charts.py`)を出す。ほかに `MasterDialog`、`ImportErrorsDialog`、`SettingsDialog`(ファイルパス・画面の大きさ・バックアップ世代数・概要で確認する日数。`[Analysis]` はメイン画面とマスタと重なるので出さない)、`constants.py`(UI文言はすべてここに定数で置く)。
- 印刷は Excel 出力(`service_excel_writer`、openpyxl のグラフ付き・A4 横 1 枚)を Excel で開いて行う。tkinter に印刷機能がないため。Excel 出力と取込エラー行の CSV 保存は `output_dir` に保存し、ファイルではなくフォルダを開く(`os.startfile`。エクスプローラーで同じフォルダが開いていれば、その窓を使う)。
- `service/service_medical_docs_analyzer.py`: `MedicalDocsAnalyzer` が統括する。取り込みは `prepare_source_import()`(元データ xlsm)/ `prepare_legacy_import()`(旧 DB xlsx)で計画(`ImportPlan`)を作り、`apply_import()` で書く。「集計開始」は元データを確認なしで取り込んでから集計し、「Excel読込」は旧 DB を確認してから取り込む。例外は握りつぶさず GUI まで送る(`NoDataError` は専用メッセージを出す)。
- `service/`: `service_excel_reader`(見出し名で列を探し、`Document` とエラー行 `RowError` に分ける)、`service_database`(SQLite の接続・スキーマ・文書の入れ替え)、`service_data_analyzer`(SQL で集計)、`service_backup`(起動時の日次バックアップ)、`service_excel_writer`(テンプレート出力)。
- `utils/`: `config_manager.py`、`date_period.py`(年・期間・年の始まりの月から開始日と終了日を求める)。
- DB は SQLite(共有フォルダに置くため `journal_mode=DELETE`)。日付は ISO 形式の文字列。院内だけで使うため患者IDは取り込む(古い DB には起動時に列を足す)。患者名・備考は取り込まない。
- 取り込みは範囲置き換え: 読んだファイルの預り日(下書きの行も含む)の最も古い日〜新しい日にある DB の文書を消して入れ直す。ファイルが古いコピーでも、それより新しい DB の文書は残る。ファイルで削除した行(取り消し)も DB から消える。
- 文書にするのは医師依頼日と担当者名がある行。空なら下書きとして飛ばす。日付が読めない・医師依頼日が預り日より前の行はエラー行として除き、一覧に出す。
- 集計対象は預り日が期間内の文書。
- 概要: 下書きは DB に入れず、`SourceData.drafts` で返す。起動時(`load_overview`、DB は開かない)・集計開始(取り込んだ計画の drafts)・設定の保存後に、預り日が本日を含め `[Overview] pending_days` 日以内の下書きを出す。状態は `draft_status`(両方空=未着手、医師依頼日だけ空=未完了、担当者名だけ空=担当不明)で、診療科ごとの件数表と行の一覧を出す。起動時に読めなくてもダイアログは出さずタブ内に出す。職員にも見せる。
- 起動時に `[Analysis] start_date / end_date` があれば `analyze_saved_period` で集計開始と同じ処理を自動で行う。
- 個人タブの文書名は全件(スクロール)、個人の Excel 出力は上位 `TOP_DOCUMENT_COUNT` 件。全体の Excel 出力は表だけで、グラフは出さない。
- 全体推移(`count_trend`): 元データは取り込まず、医師依頼日が範囲(開始日〜終了日。左側の分析期間とは連動しない)内の DB の文書を、年(年の始まりの月は左側の「1月/4月始まり」)ごとの月別に数え、1 年 1 本の折れ線にする。範囲外と本日より後の月は None(空欄)、今月は途中の件数。0 件の月は線を引かない(Excel は隠しシート「グラフ用」で 0 を空欄にする)。診療科は選択欄(先頭は合計)で選ぶ。開始日・終了日・診療科は `[Trend]` に保存し、無ければ前年(年度)の初日〜本日の合計。職員には見せない。
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
