from collections.abc import Iterator
from unittest.mock import MagicMock, patch

import pytest

from main import main


@pytest.fixture
def mock_tk() -> Iterator[MagicMock]:
    with patch('tkinter.Tk') as mock:
        yield mock


@pytest.fixture
def mock_gui() -> Iterator[MagicMock]:
    with patch('main.MedicalDocsAnalyzerGUI') as mock:
        yield mock


@pytest.fixture(autouse=True)
def mock_startup() -> Iterator[tuple[MagicMock, MagicMock, MagicMock]]:
    """本物の config.ini・ログフォルダ・出力フォルダに触れない"""
    with (patch('main.load_config') as load_config, patch('main.setup_logging') as setup_logging,
          patch('main.cleanup_old_files') as cleanup_old_files):
        yield load_config, setup_logging, cleanup_old_files


def test_main_sets_up_logging_and_cleans_output_before_login(
        mock_tk: MagicMock, mock_gui: MagicMock, mock_startup: tuple[MagicMock, MagicMock, MagicMock]) -> None:
    load_config, setup_logging, cleanup_old_files = mock_startup
    with patch('main.login_if_required', return_value=(False, None)):
        main()

    setup_logging.assert_called_once_with(load_config.return_value)
    cleanup_old_files.assert_called_once_with(load_config.return_value)


def test_main_starts_gui_after_login(mock_tk: MagicMock, mock_gui: MagicMock) -> None:
    user = MagicMock()
    with patch('main.login_if_required', return_value=(True, user)):
        main()

    mock_gui.assert_called_once_with(mock_tk.return_value, user)
    mock_tk.return_value.deiconify.assert_called_once()
    mock_tk.return_value.mainloop.assert_called_once()


def test_main_quits_when_login_cancelled(mock_tk: MagicMock, mock_gui: MagicMock) -> None:
    with patch('main.login_if_required', return_value=(False, None)):
        main()

    mock_gui.assert_not_called()
    mock_tk.return_value.destroy.assert_called_once()
    mock_tk.return_value.mainloop.assert_not_called()
