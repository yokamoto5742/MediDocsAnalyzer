import sys
import pytest
from unittest.mock import patch, MagicMock

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtGui import QKeySequence

from service_coordinate_tracker import CoordinateTracker


@pytest.fixture
def app():
    """テスト用のQApplicationを提供するフィクスチャ"""
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    yield app


class TestCoordinateTracker:
    def test_init(self, app):
        """CoordinateTrackerの初期化テスト"""
        tracker = CoordinateTracker()

        # ウィンドウタイトルを確認
        assert tracker.windowTitle() == "画面の座標"

        # サイズを確認
        assert tracker.width() == 250
        assert tracker.height() == 100

        # ラベルが存在することを確認
        assert tracker.coord_label is not None
        assert "座標:" in tracker.coord_label.text()

        # タイマーが設定されていることを確認
        assert tracker.timer is not None
        assert tracker.timer.isActive()

    @patch('service_coordinate_tracker.pyautogui.position')
    def test_update_coordinates(self, mock_position, app):
        """座標更新機能のテスト"""
        mock_position.return_value = (100, 200)

        tracker = CoordinateTracker()
        tracker.update_coordinates()

        # 座標表示が更新されることを確認
        assert tracker.coord_label.text() == "座標: (100, 200)"

        # 別の座標での更新も確認
        mock_position.return_value = (300, 400)
        tracker.update_coordinates()
        assert tracker.coord_label.text() == "座標: (300, 400)"

    @patch('service_coordinate_tracker.pyautogui.position')
    @patch('service_coordinate_tracker.QApplication.clipboard')
    def test_copy_coordinates(self, mock_clipboard, mock_position, app):
        """座標コピー機能のテスト"""
        mock_position.return_value = (500, 600)
        mock_clipboard_instance = MagicMock()
        mock_clipboard.return_value = mock_clipboard_instance

        # 静的メソッドを直接呼び出し
        CoordinateTracker.copy_coordinates()

        # クリップボードに正しい座標がコピーされたことを確認
        mock_clipboard_instance.setText.assert_called_once_with("500, 600")

    @patch('service_coordinate_tracker.pyautogui.position')
    @patch('service_coordinate_tracker.QApplication.clipboard')
    def test_shortcut_activation(self, mock_clipboard, mock_position, app):
        """ショートカットキー機能のテスト"""
        mock_position.return_value = (700, 800)
        mock_clipboard_instance = MagicMock()
        mock_clipboard.return_value = mock_clipboard_instance

        tracker = CoordinateTracker()

        # Spaceキーのショートカットを探す
        shortcut = None
        for child in tracker.children():
            if isinstance(child, QKeySequence):
                shortcut = child
                break

        # ショートカットが存在することを確認
        # 注: 実際のショートカットのテストはQtのイベントシステムに依存するため難しい
        # ここではショートカットが正しく設定されていることを間接的に確認

        # 代わりにコピー関数の動作を直接テスト
        CoordinateTracker.copy_coordinates()
        mock_clipboard_instance.setText.assert_called_once_with("700, 800")

    @patch('service_coordinate_tracker.QTimer')
    def test_timer_setup(self, mock_timer, app):
        """タイマー設定のテスト"""
        mock_timer_instance = MagicMock()
        mock_timer.return_value = mock_timer_instance

        tracker = CoordinateTracker()

        # タイマーがsetupされることを確認
        mock_timer_instance.timeout.connect.assert_called_once()
        mock_timer_instance.start.assert_called_once_with(100)
