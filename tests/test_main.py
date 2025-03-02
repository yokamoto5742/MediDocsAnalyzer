import pytest
import tkinter as tk
from unittest.mock import MagicMock, patch
from main import main


@pytest.fixture
def mock_tk():
    with patch('tkinter.Tk') as mock:
        # Tkインスタンスのモック
        mock_instance = MagicMock()
        mock.return_value = mock_instance
        yield mock


@pytest.fixture
def mock_gui():
    with patch('main.TaskAnalyzerGUI') as mock:
        yield mock


def test_main_creates_root_window(mock_tk, mock_gui):
    # mainメソッドを実行
    main()

    # Tkインスタンスが作成されたことを確認
    mock_tk.assert_called_once()

    # GUIが正しく初期化されたことを確認
    mock_gui.assert_called_once()
    assert mock_gui.call_args[0][0] == mock_tk.return_value

    # mainloopが呼び出されたことを確認
    mock_tk.return_value.mainloop.assert_called_once()


@pytest.mark.integration
def test_main_integration():
    # 統合テスト: 実際のTkinterウィンドウを使用
    with patch('tkinter.Tk.mainloop') as mock_mainloop:  # mainloopのみモック化
        main()
        mock_mainloop.assert_called_once()
