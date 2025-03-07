import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkcalendar import DateEntry
from datetime import datetime
import subprocess

from config_manager import load_config, save_config
from version import VERSION
from service_medical_docs_analyzer import MedicalDocsAnalyzer
from service_process_medical_documents import process_medical_documents


class MedicalDocsAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f'医療文書集計 v{VERSION}')  # タイトルを「医療文書集計」に変更
        self.config = load_config()
        self.analyzer = MedicalDocsAnalyzer()

        window_width = self.config.getint('Appearance', 'window_width')
        window_height = self.config.getint('Appearance', 'window_height')
        self.root.geometry(f"{window_width}x{window_height}")

        self._setup_gui()

    def _setup_gui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self._setup_date_frame(main_frame)
        self._setup_buttons(main_frame)

    def _setup_date_frame(self, parent):
        date_frame = ttk.LabelFrame(parent, text="分析期間", padding="5")
        date_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(date_frame, text="開始日").grid(row=0, column=0, padx=5, pady=5)

        # 開始日の初期値をコンフィグから取得
        start_date_str = self.config.get('Analysis', 'start_date', fallback='2025-01-01')
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d')

        self.start_date = DateEntry(date_frame, width=12, background='darkblue',
                                    foreground='white', borderwidth=2,
                                    year=start_date.year, month=start_date.month, day=start_date.day,
                                    locale='ja_JP', date_pattern='yyyy/mm/dd')
        self.start_date.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(date_frame, text="終了日:").grid(row=1, column=0, padx=5, pady=5)

        # 終了日の初期値をコンフィグから取得
        end_date_str = self.config.get('Analysis', 'end_date', fallback='2025-01-31')
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d')

        self.end_date = DateEntry(date_frame, width=12, background='darkblue',
                                  foreground='white', borderwidth=2,
                                  year=end_date.year, month=end_date.month, day=end_date.day,
                                  locale='ja_JP', date_pattern='yyyy/mm/dd')
        self.end_date.grid(row=1, column=1, padx=5, pady=5)

    def _setup_buttons(self, parent):
        # データ読込ボタンを追加
        ttk.Button(parent, text="データ読込", command=self.load_data).grid(
            row=2, column=0, columnspan=2, pady=10)

        ttk.Button(parent, text="集計開始", command=self.start_analysis).grid(
            row=3, column=0, columnspan=2, pady=10)
        ttk.Button(parent, text="設定ファイル", command=self.open_config).grid(
            row=4, column=0, columnspan=2, pady=5)
        ttk.Button(parent, text="閉じる", command=self.root.quit).grid(
            row=5, column=0, columnspan=2, pady=5)

    # データ読込機能を追加
    def load_data(self):
        """GUIからデータ読込を開始するメソッド"""
        try:
            # 設定ファイルからパスを取得
            source_file_path = self.config.get('PATHS', 'source_file_path')
            database_path = self.config.get('PATHS', 'database_path')

            # 日付をコンフィグに保存
            self.save_date_to_config()

            # データ処理を実行
            success = process_medical_documents(source_file_path, database_path)

            if success:
                messagebox.showinfo("成功", "データの読込が完了しました。")
            else:
                messagebox.showerror("エラー", "データの読込中にエラーが発生しました。")

        except Exception as e:
            messagebox.showerror("エラー", f"データ読込中に予期せぬエラーが発生しました：\n{str(e)}")

    def save_date_to_config(self):
        """日付設定をコンフィグに保存するメソッド"""
        start_date = self.start_date.get_date()
        end_date = self.end_date.get_date()

        if start_date > end_date:
            messagebox.showerror("エラー", "開始日が終了日より後の日付になっています。")
            return False

        # 設定の保存
        if 'Analysis' not in self.config:
            self.config.add_section('Analysis')

        self.config['Analysis'].update({
            'start_date': start_date.strftime('%Y-%m-%d'),
            'end_date': end_date.strftime('%Y-%m-%d')
        })
        save_config(self.config)
        return True

    def start_analysis(self):
        """GUIから分析を開始するメソッド"""
        try:
            # 日付をコンフィグに保存
            if not self.save_date_to_config():
                return

            start_date = self.start_date.get_date()
            end_date = self.end_date.get_date()

            # 分析の実行
            success, error_message = self.analyzer.run_analysis(
                start_date.strftime('%Y-%m-%d'),
                end_date.strftime('%Y-%m-%d')
            )

            if not success:
                messagebox.showerror("エラー", error_message)
            else:
                messagebox.showinfo("成功", error_message)

        except Exception as e:
            messagebox.showerror("エラー", f"予期せぬエラーが発生しました：\n{str(e)}")

    def open_config(self):
        config_path = self.config.get('PATHS', 'config_path')

        try:
            subprocess.Popen(['notepad.exe', config_path])
        except Exception as e:
            messagebox.showerror("エラー", f"設定ファイルを開けませんでした：\n{str(e)}")