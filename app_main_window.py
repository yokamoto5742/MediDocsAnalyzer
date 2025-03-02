import os
import sys
from pathlib import Path

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIntValidator
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QPushButton, QLabel, QMessageBox
)

from app_dialogs import ExcludeDocsDialog, ExcludeDoctorsDialog, AppearanceDialog, FolderPathDialog
from config_manager import ConfigManager
from service_coordinate_tracker import CoordinateTracker
from service_csv_excel_transfer import transfer_csv_to_excel
from version import VERSION


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = ConfigManager()
        self.tracker = CoordinateTracker()
        font = self.font()
        font.setPointSize(self.config.get_font_size())
        self.setFont(font)
        window_size = self.config.get_window_size()
        self.setFixedSize(*window_size)
        self.setWindowTitle(f"CSV取込アプリ v{VERSION}")

        # メインウィジェット
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        layout = QVBoxLayout()

        self.setStyleSheet("QMainWindow { border: 5px solid darkgreen; }")

        title_label = QLabel("Papyrus書類受付リスト")
        layout.addWidget(title_label)

        csv_button = QPushButton("CSVファイル取り込み")
        csv_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 8px;
                border: 2px solid #45a049;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """)
        csv_button.clicked.connect(self.import_csv)
        layout.addWidget(csv_button)

        settings_label = QLabel("設定")
        layout.addWidget(settings_label)

        exclude_docs_button = QPushButton("除外する文書名")
        exclude_docs_button.clicked.connect(self.show_exclude_docs_dialog)
        layout.addWidget(exclude_docs_button)

        exclude_doctors_button = QPushButton("除外する医師名")
        exclude_doctors_button.clicked.connect(self.show_exclude_doctors_dialog)
        layout.addWidget(exclude_doctors_button)

        appearance_button = QPushButton("フォントとウインドウサイズ")
        appearance_button.clicked.connect(self.show_appearance_dialog)
        layout.addWidget(appearance_button)

        coordinate_button = QPushButton("画面の座標表示")
        coordinate_button.clicked.connect(self.show_coordinate_tracker)
        layout.addWidget(coordinate_button)

        folder_path_button = QPushButton("フォルダの場所")
        folder_path_button.clicked.connect(self.show_folder_path_dialog)
        layout.addWidget(folder_path_button)

        close_button = QPushButton("閉じる")
        close_button.clicked.connect(self.close)
        layout.addWidget(close_button)

        main_widget.setLayout(layout)

    def import_csv(self):
        try:
            transfer_csv_to_excel()
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"CSVファイルの取り込み中にエラーが発生しました:\n{str(e)}")

    def show_exclude_docs_dialog(self):
        dialog = ExcludeDocsDialog(self)
        dialog.exec()

    def show_exclude_doctors_dialog(self):
        dialog = ExcludeDoctorsDialog(self)
        dialog.exec()

    def show_appearance_dialog(self):
        dialog = AppearanceDialog(self)
        dialog.exec()

    def show_coordinate_tracker(self):
        self.tracker.show()

    def show_folder_path_dialog(self):
        dialog = FolderPathDialog(self)
        dialog.exec()
