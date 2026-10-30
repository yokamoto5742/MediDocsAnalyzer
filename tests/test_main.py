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
