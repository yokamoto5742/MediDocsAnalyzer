import os
import sys
from pathlib import Path

from PyQt6.QtGui import QIntValidator
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QLineEdit, QListWidget, QDialogButtonBox, QFileDialog,
    QMessageBox
)

from config_manager import ConfigManager

class ExcludeItemDialog(QDialog):
    def __init__(self, title, item_label, config_section, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.config_section = config_section

        layout = QVBoxLayout()

        self.input_field = QLineEdit()
        layout.addWidget(QLabel(f"{item_label}を入力:"))
        layout.addWidget(self.input_field)

        self.item_list = QListWidget()
        layout.addWidget(QLabel(f"登録済み{item_label}:"))
        layout.addWidget(self.item_list)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        add_button = QPushButton("追加")
        add_button.clicked.connect(self.add_item)

        delete_button = QPushButton("削除")
        delete_button.clicked.connect(self.delete_selected)

        layout.addWidget(add_button)
        layout.addWidget(delete_button)
        layout.addWidget(buttons)

        self.setLayout(layout)

        self.config = ConfigManager()
        self.load_items()

    def load_items(self):
        if self.config_section in self.config.config:
            items = self.config.config[self.config_section].get('list', '').split(',')
            for item in items:
                if item.strip():
                    self.item_list.addItem(item.strip())

    def add_item(self):
        item_name = self.input_field.text().strip()
        if item_name:
            self.item_list.addItem(item_name)
            self.input_field.clear()

    def delete_selected(self):
        current_item = self.item_list.currentItem()
        if current_item:
            self.item_list.takeItem(self.item_list.row(current_item))

    def accept(self):
        items = []
        for i in range(self.item_list.count()):
            items.append(self.item_list.item(i).text())

        if self.config_section not in self.config.config:
            self.config.config[self.config_section] = {}
        self.config.config[self.config_section]['list'] = ','.join(items)
        self.config.save_config()
        super().accept()


class ExcludeDocsDialog(ExcludeItemDialog):
    def __init__(self, parent=None):
        super().__init__("除外する文書名", "除外する文書名", "ExcludeDocs", parent)


class ExcludeDoctorsDialog(ExcludeItemDialog):
    def __init__(self, parent=None):
        super().__init__("除外する医師名", "除外する医師名", "ExcludeDoctors", parent)


class FolderPathDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("フォルダの場所")
        self.setModal(True)

        config = ConfigManager()
        dialog_width, dialog_height = config.get_folder_dialog_size()
        self.resize(dialog_width, dialog_height)

        layout = QVBoxLayout()

        self.downloads_path = self.create_path_section(
            layout, "ダウンロードフォルダ:",
            lambda: self.browse_folder('downloads'),
            lambda: self.open_folder(self.downloads_path.text())
        )

        self.excel_path = self.create_path_section(
            layout, "Excelファイルパス:",
            lambda: self.browse_folder('excel'),
            lambda: self.open_folder(str(Path(self.excel_path.text()).parent))
        )

        self.backup_path = self.create_path_section(
            layout, "バックアップフォルダ:",
            lambda: self.browse_folder('backup'),
            lambda: self.open_folder(self.backup_path.text())
        )

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

        self.config = ConfigManager()
        self.load_paths()

    @staticmethod
    def create_path_section(layout, label_text, browse_callback, open_callback):
        layout.addWidget(QLabel(label_text))
        path_layout = QHBoxLayout()

        path_field = QLineEdit()
        path_layout.addWidget(path_field)

        browse_button = QPushButton("参照...")
        browse_button.clicked.connect(browse_callback)
        path_layout.addWidget(browse_button)

        open_button = QPushButton("開く")
        open_button.clicked.connect(open_callback)
        path_layout.addWidget(open_button)

        layout.addLayout(path_layout)
        return path_field

    def open_folder(self, path: str):
        if path and os.path.exists(path):
            os.startfile(path)
        else:
            QMessageBox.warning(self, "警告", "指定されたフォルダが存在しません。")

    def load_paths(self):
        self.downloads_path.setText(self.config.get_downloads_path())
        self.excel_path.setText(self.config.get_excel_path())
        self.backup_path.setText(self.config.get_backup_path())

    def browse_folder(self, path_type):
        if path_type in ['downloads', 'backup']:
            title = "ダウンロードフォルダの選択" if path_type == 'downloads' else "バックアップフォルダの選択"
            path_field = self.downloads_path if path_type == 'downloads' else self.backup_path

            folder = QFileDialog.getExistingDirectory(
                self,
                title,
                path_field.text()
            )
            if folder:
                path_field.setText(folder)
        else:
            file, _ = QFileDialog.getOpenFileName(
                self,
                "Excelファイルの選択",
                self.excel_path.text(),
                "Excel Files (*.xlsx *.xlsm)"
            )
            if file:
                self.excel_path.setText(file)

    def accept(self):
        self.config.set_downloads_path(self.downloads_path.text())
        self.config.set_excel_path(self.excel_path.text())
        self.config.set_backup_path(self.backup_path.text())
        super().accept()


class AppearanceDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("フォントとウインドウサイズ")
        self.setModal(True)

        layout = QVBoxLayout()

        layout.addWidget(QLabel("フォントサイズ:"))
        self.font_size_input = QLineEdit()
        self.font_size_input.setValidator(QIntValidator(6, 72))
        layout.addWidget(self.font_size_input)

        layout.addWidget(QLabel("ウィンドウの幅:"))
        self.window_width_input = QLineEdit()
        self.window_width_input.setValidator(QIntValidator(200, 1000))
        layout.addWidget(self.window_width_input)

        layout.addWidget(QLabel("ウィンドウの高さ:"))
        self.window_height_input = QLineEdit()
        self.window_height_input.setValidator(QIntValidator(150, 800))
        layout.addWidget(self.window_height_input)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

        self.config = ConfigManager()
        self.load_settings()

    def load_settings(self):
        self.font_size_input.setText(str(self.config.get_font_size()))
        window_size = self.config.get_window_size()
        self.window_width_input.setText(str(window_size[0]))
        self.window_height_input.setText(str(window_size[1]))

    def accept(self):
        self.config.set_font_size(int(self.font_size_input.text()))
        self.config.set_window_size(
            int(self.window_width_input.text()),
            int(self.window_height_input.text())
        )
        QMessageBox.information(self, "設定完了", "設定を保存しました。\n変更を適用するにはアプリケーションを再起動してください。")
        super().accept()
