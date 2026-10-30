import tkinter as tk
from collections.abc import Iterator
from unittest.mock import MagicMock, patch

import pytest

from app import constants
from app.login_dialog import ChangePasswordDialog, LoginDialog, ask_login, login_if_required
from service.service_auth import LoginResult
from service.service_master import Staff

USER = Staff(1, '花﨑', 1, emr_id='S01')


@pytest.fixture
def messagebox() -> Iterator[MagicMock]:
    with patch('app.login_dialog.messagebox') as mock:
        yield mock


def test_login_succeeds(root: tk.Tk, messagebox: MagicMock) -> None:
    analyzer = MagicMock()
    analyzer.authenticate.return_value = LoginResult(USER, must_change_password=True)
    dialog = LoginDialog(root, analyzer)
    dialog.emr_id_var.set(' S01 ')
    dialog.password_var.set('secret1')

    dialog.login()

    analyzer.authenticate.assert_called_once_with('S01', 'secret1')
    assert (dialog.user, dialog.must_change_password) == (USER, True)


def test_login_fails(root: tk.Tk, messagebox: MagicMock) -> None:
    analyzer = MagicMock()
    analyzer.authenticate.return_value = None
    dialog = LoginDialog(root, analyzer)
    dialog.password_var.set('wrong')

    dialog.login()

    assert dialog.user is None
    assert dialog.password_var.get() == ''
    assert messagebox.showerror.call_args.args == (constants.ERROR_TITLE, constants.LOGIN_FAILED_MESSAGE)


@pytest.mark.parametrize(('password', 'confirm', 'expected_message'), [
    ('abc', 'abc', constants.PASSWORD_TOO_SHORT_MESSAGE.format(length=6)),
    ('secret1', 'secret2', constants.PASSWORD_MISMATCH_MESSAGE),
])
def test_change_password_validates(root: tk.Tk, messagebox: MagicMock,
                                   password: str, confirm: str, expected_message: str) -> None:
    analyzer = MagicMock()
    dialog = ChangePasswordDialog(root, analyzer, USER)
    dialog.password_var.set(password)
    dialog.confirm_var.set(confirm)

    dialog.save()

    assert messagebox.showerror.call_args.args == (constants.ERROR_TITLE, expected_message)
    analyzer.change_password.assert_not_called()
    assert not dialog.changed


def test_change_password_saves(root: tk.Tk, messagebox: MagicMock) -> None:
    analyzer = MagicMock()
    dialog = ChangePasswordDialog(root, analyzer, USER)
    dialog.password_var.set('secret1')
    dialog.confirm_var.set('secret1')

    dialog.save()

    analyzer.change_password.assert_called_once_with(1, 'secret1')
    assert dialog.changed


def test_ask_login_requires_password_change(root: tk.Tk, messagebox: MagicMock) -> None:
    login_dialog = MagicMock(user=USER, must_change_password=True)
    change_dialog = MagicMock(changed=False)
    with (patch('app.login_dialog.LoginDialog', return_value=login_dialog),
          patch('app.login_dialog.ChangePasswordDialog', return_value=change_dialog),
          patch.object(root, 'wait_window')):
        assert ask_login(root, MagicMock()) is None
        change_dialog.changed = True
        assert ask_login(root, MagicMock()) == USER


@pytest.mark.parametrize(('required', 'user', 'expected'), [
    (False, None, (True, None)),
    (True, USER, (True, USER)),
    (True, None, (False, None)),
])
def test_login_if_required(root: tk.Tk, messagebox: MagicMock, required: bool, user: Staff | None,
                           expected: tuple[bool, Staff | None]) -> None:
    with (patch('app.login_dialog.MedicalDocsAnalyzer') as analyzer_class,
          patch('app.login_dialog.ask_login', return_value=user)):
        analyzer_class.return_value.login_required.return_value = required
        assert login_if_required(root) == expected


def test_login_if_required_starts_without_login_when_database_fails(root: tk.Tk, messagebox: MagicMock) -> None:
    with patch('app.login_dialog.MedicalDocsAnalyzer') as analyzer_class:
        analyzer_class.return_value.login_required.side_effect = OSError('offline')
        assert login_if_required(root) == (True, None)

    messagebox.showerror.assert_called_once()
