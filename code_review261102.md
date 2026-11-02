# コードレビュー (2026-11-02, v1.13.1)

対象: `main.py`、`app/`、`service/`、`utils/` の全ソース
観点: 可読性とメンテナンス性、KISS(できるだけシンプルにする)

## 総評

層の分け方(GUI → `MedicalDocsAnalyzer` → 各 service → SQLite / openpyxl)がはっきりしていて、SQL も短く読みやすいです。型ヒント、日本語の短い docstring、「なぜそうするか」だけを書いたコメントも規約どおりです。

一方で、機能を足すたびに **同じ考え方が 2〜3 か所に書かれる** 箇所が増えてきています。中でも次の 3 つは、片方だけ直すと表示と Excel、または GUI とサービスで結果が食い違う原因になります。

1. 「合計」という文字列が、サービスの判定と UI の表示の両方で使われている
2. 全体集計の表が、GUI(`overall_rows`)と Excel(`write_counts`)で別々に組み立てられている
3. 設定を「読み直して → 更新して → 保存」する処理が 3 か所にある

下では優先度(高・中・低)ごとに、指摘と直し方の案を書きます。

---

## 優先度: 高

### H1. 「合計」が UI の文言とサービスの判定に二重に定義されている

- `service/service_excel_writer.py:16` `TOTAL_LABEL = '合計'`
- `app/constants.py:39` `TOTAL_ROW_LABEL = "合計"`
- `app/app_window.py:291` は診療科の選択肢の先頭に `constants.TOTAL_ROW_LABEL` を入れる
- `service/service_medical_docs_analyzer.py:230` は `department == TOTAL_LABEL` のとき全診療科と判定する

UI の文言を「全科」などに変えると、推移の「合計」が普通の診療科名として SQL に渡され、0 件になります。また `TOTAL_LABEL` が Excel 出力のモジュールにあるため、`service_medical_docs_analyzer` が「合計」を知るためだけに writer を import しています(`:18`)。

**案:** 「全診療科」を文字列ではなく `None` で表す。

```python
# app_window.py: 選択肢の表示は UI 側で、サービスには None を渡す
department = panel.selected_department()
department_filter = None if department == constants.TOTAL_ROW_LABEL else department
self.analyzer.trend_summary(department_filter, ...)

# service_medical_docs_analyzer.py
def trend_summary(self, department: str | None, ...) -> TrendCounts:
    with self._open_database() as connection:
        return count_trend(connection, department, ...)
```

`count_trend` は最初から `department: str | None` を受け取るので、変換している 1 行が消えるだけです。`seed_masters` の前に「合計」を除いている処理(`:107-108`)は、古い ini との互換のため残します。

### H2. 全体集計の表を GUI と Excel で別々に組み立てている

- `app/result_view.py:37-46` `overall_rows()`(docstring に「Excel 出力と同じ行」とある)
- `service/service_excel_writer.py:20-48` `write_counts()`

同じ表(担当者 × 診療科、合計列、合計行)を 2 回実装しています。`write_counts` は「合計」列がどこにあっても動くように作られているため、29 行あって読みにくくなっています。実際には合計列はいつも先頭です(`OverallSummary.departments` の「先頭は合計列」という約束、`:97`)。

この約束のため、ほかの場所にも `summary.departments[1:]` が出てきます(`app/charts.py:24`、`app/result_view.py:40`)。

**案:**

1. `OverallSummary.departments` には診療科だけを入れ、合計列は表示するときに足す。
2. 行の組み立てを `OverallSummary` のプロパティ 1 つにまとめ、GUI と Excel の両方で使う。

```python
@dataclass
class OverallSummary:
    counts: DocumentCounts
    staff_members: list[str]
    departments: list[str]          # 診療科だけ

    @property
    def rows(self) -> list[list[str | int]]:
        """[氏名, 合計, 診療科ごとの件数...]。最終行の合計は一覧にない診療科も含めた全体の件数"""
        ...

# service_excel_writer.py
def write_counts(sheet: Worksheet, summary: OverallSummary) -> None:
    for column_index, header in enumerate(['氏名', TOTAL_LABEL, *summary.departments], start=1):
        sheet.cell(row=2, column=column_index, value=header)
    for row_index, row in enumerate(summary.rows, start=3):
        for column_index, value in enumerate(row, start=1):
            sheet.cell(row=row_index, column=column_index, value=value)
```

これで `write_counts` は 10 行以下になり、`if department == TOTAL_LABEL: continue` や `departments.index(TOTAL_LABEL)` も要らなくなります。`save_results` の引数も `(summary, template_path, output_dir)` に減ります。

### H3. GUI の外で起きた例外が、利用者に見えずに消える

Tk のボタンから呼ばれる処理で起きた例外は、呼び出し元の `try` には戻らず、Tk の `report_callback_exception` で stderr に出るだけです。PyInstaller の windowed exe では何も表示されません。

| 場所 | 起きること |
|---|---|
| `app/app_window.py:378-387` `change_password` | `try` で囲んでいるのはダイアログを作る処理だけ。「保存」ボタン(`ChangePasswordDialog.save`)の DB エラーはここに来ない。さらにメッセージが `ANALYSIS_ERROR_MESSAGE`(「集計中にエラー」)になっている |
| `app/login_dialog.py:46-54` `LoginDialog.login` | DB に接続できないと、ログインボタンを押しても何も起きない |
| `app/master_dialog.py:176-190` `issue_temporary_password` | 仮パスワードを発行できなくても何も出ない |
| `app/master_dialog.py:209-217` `delete_department` | 同上 |

**案:** DB を使うコールバックの中で、例外をメッセージに変える。メッセージは処理に合ったものにする(`PASSWORD_CHANGE_ERROR_MESSAGE` などを `constants.py` に足す)。どこか 1 か所でまとめて拾うなら、`root.report_callback_exception` を上書きして `messagebox.showerror` とログを出すだけでも、見えない失敗はなくなります。

```python
# main.py
def report_error(exc_type, exc_value, traceback) -> None:
    logging.error('想定外のエラー', exc_info=(exc_type, exc_value, traceback))
    messagebox.showerror(constants.ERROR_TITLE, str(exc_value))

root.report_callback_exception = report_error
```

---

## 優先度: 中

### M1. 設定の「読み直し → 更新 → 保存」が 3 か所にある

- `app/app_window.py:355-366` `_save_trend_settings`
- `app/app_window.py:368-376` `_save_analysis_settings`
- `app/settings_dialog.py:108-115` `save`

3 か所とも「ほかの画面の変更を消さないよう、保存直前に読み直す」という同じ理由で同じ手順を書いています。さらに `app_window` は読み直した設定を `self.config` に入れ直しています。そのため `self.config` が「起動時の設定」なのか「最後に保存した設定」なのか、読んだだけでは分かりません。

**案:** `utils/config_manager.py` に 1 つ関数を置く。

```python
def update_section(section: str, values: dict[str, str], replace: bool = False) -> None:
    """ほかの画面の変更を消さないよう、保存直前に読み直してから section を書き換える"""
    config = load_config()
    if replace or not config.has_section(section):
        config[section] = {}
    config[section].update(values)
    save_config(config)
```

`self.config` は `__init__` で読むだけにして、保存では使わないようにします。

### M2. 全体推移と個人推移の「入力チェック」が、表示と出力で別々に書かれている

`app/app_window.py:317-334` `_refresh_trend_panel` と `:342-353` `_export_trend_panel` は、どちらも「診療科を取る → 氏名を取る(`''` なら中止)→ 範囲を読む」をしています。条件の書き方が少し違うため(片方は `trend_range` を先に読む)、比べないと同じ処理だと分かりません。

また `selected_staff()` が返すのは次の 3 通りで、`staff_name == ''` という比較の意味がコメントを読まないと分かりません。

- `None`: 全体推移
- `''`: 担当者がいない
- 氏名

**案:** 入力をまとめて読む関数を 1 つにする。

```python
@dataclass
class TrendRequest:
    department: str | None   # None なら全診療科(H1)
    staff_name: str | None   # None なら全担当者
    start_date: date
    end_date: date

def _read_trend_request(self, panel: TrendPanel) -> TrendRequest | None:
    """推移の選択と範囲。個人推移で担当者がいないか、範囲が正しくなければ None"""
```

表示側と出力側は `request = self._read_trend_request(panel)` の 1 行で始められます。

### M3. 日付入力欄の作成と日付範囲の読み取りが重複している

- `DateEntry(..., width=12, background='darkblue', foreground='white', borderwidth=2, locale='ja_JP', date_pattern='yyyy/mm/dd')` が `app/app_window.py:86-88` と `app/result_view.py:124-125` にある
- 「`parse_date(entry.get())` を 2 回呼び、開始 ≤ 終了を確かめる」処理が `app/app_window.py:203-212` と `:302-309` にある

**案:** `app/` に `make_date_entry(parent)` と `read_date_range(start_entry, end_entry) -> tuple[date, date] | None` を置き、両方から使う。`get_date()` ではなく文字列を読む理由のコメント(`:202`)も、その関数に 1 回だけ書けば済みます。

### M4. `ImportPlan` が assert 付きのプロパティで日付を返している

`service/service_medical_docs_analyzer.py:72-90`: `start_date` と `end_date` は `source.first_received_date` が `None` でないことを assert で確かめています。ですが、`_prepare_import`(`:130-131`)で `None` なら `NoDataError` にしているので、`ImportPlan` ができた時点で日付は必ずあります。

**案:** 作るときに値を渡す。

```python
@dataclass
class ImportPlan:
    source: SourceData
    start_date: date
    end_date: date
    deleted_count: int
```

プロパティ 3 つと assert 2 つが消えます。`errors` プロパティも、呼び出し元で `plan.source.errors` と書けば要りません。

### M5. Excel の 1 行で、同じセルの日付を 2〜3 回解釈している

`service/service_excel_reader.py`:

- `_add_row` で預り日を `to_date` する(`:133-136`)
- `row_error_reason` で預り日と医師依頼日をもう一度 `to_date` する(`:84-93`)
- 正常な行では、さらに医師依頼日を `to_date` する(`:147`)

結果は同じですが、「この時点で `received` は `None` ではない」という前提を `assert`(`:149`)で補っていて、流れを追いにくくなっています。

**案:** 1 回だけ解釈して、値かエラー理由を返す。

```python
def parse_row(cells: dict[str, object], known_staff: set[str] | None) -> tuple[date, date | None, str] | str:
    """(預り日, 医師依頼日, 担当者名) を返す。取り込めない行はエラーの理由(文字列)を返す"""
```

あわせて、`'文書名'`、`'診療科'`、`'医師名'` だけが定数になっていない点(`:14`、`:150`)もそろえると、列名の一覧が 1 か所で分かります。

### M6. サービス層がエクスプローラーを開いている

`service/service_excel_writer.py:67-77` `save_and_open` は、保存した後に `os.startfile` で出力先フォルダを開きます。画面に関わる処理がサービスに入っているため、writer のテストでは `startfile` をモックする必要があります。また、CSV 保存(`app/import_errors_dialog.py:71`)にも同じ `startfile` があります。

**案:** writer は保存してパスを返すだけにします。フォルダを開くのは GUI 側の小さな関数 `open_folder(path)` にまとめ、Excel 出力と CSV 保存の両方から呼びます。

### M7. 職員(ロール)の制限が 2 か所に分かれている

- `app/app_window.py:62-68` `_apply_permissions` が `self.user.role` で判定し、ボタンを無効にする
- 一方で `:271` と `:297` は `self.result_view.restricted_staff` を見て判定する
- `app/result_view.py:301-306` `set_staff_names` も、内部で `restricted_staff` に置き換える

「職員かどうか」を知っている場所が GUI(`self.user`)と ResultView(`restricted_staff`)の 2 つあります。

**案:** 判定を `MedicalDocsAnalyzerGUI` のプロパティ 1 つ(例: `restricted_staff -> str | None`)にまとめる。ResultView は渡された名前の一覧を表示するだけにして、「本人だけにする」判断は GUI 側で行う。

### M8. `ResultView` の引数が 7 つのコールバックで、順番だけで区別されている

`app/result_view.py:169-175` と、それを呼ぶ `app/app_window.py:56-58` は、`Callable[[], None]` が 7 つ並んでいます。型では区別できないので、順番を入れ替えても Pyright は気づきません。

**案:** 呼び出し側でキーワード引数にする(コードの変更はほぼ不要)。

```python
self.result_view = ResultView(
    self.root,
    on_export_overall=self.export_overall,
    on_export_staff=self.export_staff,
    ...
)
```

また、`__init__` の中でタブを作る順と `add` する順が混ざっています(`:181-190`)。直前のコミットでタブを並べ替えた影響です。「作ってすぐ `add`」を表示順に並べると、並びが一目で分かります。

---

## 優先度: 低

### L1. UI の文言が `constants.py` に移っていない

CLAUDE.md の「UI文言はすべて `constants.py` に定数で置く」に合っていない箇所です。

- `app/app_window.py:71` `"分析期間"`、`:80-81` `"開始日"` `"終了日"`、`:101` `"年"`、`:110` `"{month}月始まり"`、`:114` `"期間"`、`:134` `"集計開始"`、`:145` `"閉じる"`(`CLOSE_BUTTON_LABEL` が既にある)
- `app/charts.py:32,37,50,68` `'件数'`、`'月別件数'`

### L2. `fiscal_year` を import しているのに、同じ計算を手で書いている

`app/app_window.py:76-77`

```python
# 4月始まりなら 1〜3月は前年の年度として表示する
initial_year = saved_start_date.year - (saved_start_date.month < year_start_month)
```

→ `initial_year = fiscal_year(saved_start_date, year_start_month)` の 1 行にでき、コメントも要らなくなります(`fiscal_year` の docstring に同じ説明がある)。

### L3. 期間の条件 `received_date BETWEEN ? AND ?` が散らばっている

`service/service_data_analyzer.py:9,54` には `PERIOD_CONDITION` と `period_parameters` がありますが、`service/service_database.py:74,81` には同じ SQL と `isoformat()` を直接書いています。どちらかに寄せると、預り日の条件を変えるときに 1 か所で済みます。

### L4. 取り込みのエラー処理が 2 つのメソッドで同じ

`app/app_window.py:164-178` `load_excel` と `:186-199` `_import_source` は、どちらも次の 3 つを持っています。

- `NoDataError` → `NO_SOURCE_DATA_MESSAGE`
- `Exception` → `LOAD_ERROR_MESSAGE`
- 取り込めない行があればダイアログを出す

`_import_source(confirm: bool)` の 1 つにするか、少なくとも「エラー行があればダイアログを出す」部分を共通にできます。

### L5. DB の行から `Staff` を作る処理が 2 か所にある

`service/service_auth.py:46-54` と `service/service_master.py:34-38` は、同じ列から `Staff(...)` を組み立てています。`service_auth.authenticate` ではタプルを展開するとき、引数の `emr_id` を同じ名前の変数で上書きしています。`service_master` に `STAFF_COLUMNS` と `_staff_from_row(row)` を置き、auth は追加の 3 列だけを別に読むと、列を追加するときに 1 か所の修正で済みます。

### L6. 共通のデータ型が `service_excel_reader` に置かれている

`Document` と `NoDataError` は reader で定義されていて、`service_database` と `service_data_analyzer` は reader を import してそれを使っています。DB や集計が「Excel の読み込み」に依存する形になっています。今の規模なら困りませんが、データ型が増えるなら `service/models.py` に移すと依存の向きが自然になります。

同じように、`service_medical_docs_analyzer.py` には次の 3 つの役割が同居しています。

- 概要(`DraftStatus`、`Overview`、`make_overview`)
- 取り込み計画(`ImportPlan`)
- 窓口クラス(`MedicalDocsAnalyzer`)

概要は `service_overview.py` に分けると、250 行のファイルが半分ほどになります。

### L7. ログの古いファイルを 2 つの仕組みで消している

`utils/log_rotation.py` は `TimedRotatingFileHandler(backupCount=...)` で世代を管理したうえで、`cleanup_old_logs` でも古い `.log` を消しています。どちらか一方にすると、保持期間の考え方が 1 つになります。また、`utils/file_cleaner.py` だけが Google 形式の docstring(`Args:`)を使っていて、ほかのファイルの書き方と違います。ログでは `f"{str(e)}"` の `str()` も不要です。

### L8. GUI のコンストラクタが重い

`MedicalDocsAnalyzerGUI.__init__`(`app/app_window.py:26-45`)は、画面を作るだけでなく、次の処理まで行います。

- バックアップ
- 元データの読み込み
- 推移の集計
- 取り込みと集計

起動時の処理を `start()` などに分けて `main.py` から呼ぶと、テストで画面だけを作りやすくなり、「起動時に何が起きるか」も `main.py` を見れば分かります。

---

## 直す順番の提案

| 順 | 項目 | 効果 | 規模 |
|---|---|---|---|
| 1 | H3 例外が見えない | 不具合の見落としを防ぐ | 小 |
| 2 | H1 「合計」を `None` に | UI 文言を変えても壊れない | 小 |
| 3 | L1, L2 文言の定数化・`fiscal_year` | 規約に合わせる | 小 |
| 4 | M1 設定保存の共通化 | 3 か所 → 1 か所 | 小 |
| 5 | H2 全体表の一本化 | 表示と Excel のずれを防ぐ。`write_counts` が約 1/3 に | 中 |
| 6 | M2, M3 推移・日付入力の共通化 | `app_window.py` が約 40 行減る | 中 |
| 7 | M4, M5 取り込み周りの整理 | assert と同じ値の再解釈が消える | 中 |
| 8 | M6〜M8, L3〜L8 | 依存関係と見通しの改善 | 適宜 |

どれも振る舞いを変えないリファクタリングです(H3 だけは「エラーを表示する」という変更を含みます)。既存のテスト(`tests/app/test_app_window.py`、`tests/service/test_service_excel_writer.py` など)が通ることを、変更ごとに確認しながら進めてください。
