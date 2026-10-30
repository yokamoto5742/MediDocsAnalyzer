import tkinter as tk
from collections.abc import Iterator
from pathlib import Path
from tkinter import ttk
from unittest.mock import MagicMock, patch

import pytest

from app import constants
from app.master_dialog import MasterDialog
from service.service_master import ADMIN_ROLE, Department, Staff
from service.service_medical_docs_analyzer import MedicalDocsAnalyzer


@pytest.fixture
def messagebox() -> Iterator[MagicMock]:
    with patch('app.master_dialog.messagebox') as mock:
        yield mock


@pytest.fixture
def analyzer(config_path: Path) -> MedicalDocsAnalyzer:
    return MedicalDocsAnalyzer()


@pytest.fixture
def dialog(root: tk.Tk, analyzer: MedicalDocsAnalyzer, messagebox: MagicMock) -> MasterDialog:
    return MasterDialog(root, analyzer)


def tree_values(tree: ttk.Treeview) -> list[tuple[str, ...]]:
    return [tuple(tree.item(item, 'values')) for item in tree.get_children()]


def test_init_lists_seeded_masters(dialog: MasterDialog) -> None:
    assert tree_values(dialog.staff_tree) == [
        ('植田', '1', '在籍', '', '職員'), ('沖野', '2', '在籍', '', '職員'), ('花﨑', '3', '在籍', '', '職員')]
    assert tree_values(dialog.department_tree) == [('内科', '1'), ('外科', '2')]
    assert dialog.staff_order_var.get() == '4'


def test_save_staff_updates_selected(dialog: MasterDialog, analyzer: MedicalDocsAnalyzer) -> None:
    dialog.staff_tree.selection_set('1')
    dialog._on_staff_selected()
    dialog.staff_active_var.set(False)
    dialog.staff_emr_id_var.set(' E02 ')
    dialog.staff_role_var.set(constants.ADMIN_ROLE_LABEL)

    dialog.save_staff()

    okino = analyzer.list_staff()[1]
    assert (okino.name, okino.active, okino.emr_id, okino.role) == ('沖野', False, 'E02', ADMIN_ROLE)


def test_save_staff_adds_new(dialog: MasterDialog, analyzer: MedicalDocsAnalyzer) -> None:
    dialog.staff_name_var.set('小牧')

    dialog.save_staff()

    assert analyzer.list_staff()[-1] == Staff(4, '小牧', 4)


@pytest.mark.parametrize(('name', 'order', 'expected_message'), [
    ('', '1', constants.NAME_REQUIRED_MESSAGE),
    ('小牧', 'a', constants.INVALID_ORDER_MESSAGE),
    ('植田', '1', constants.DUPLICATE_STAFF_MESSAGE),
])
def test_save_staff_shows_error(dialog: MasterDialog, messagebox: MagicMock,
                                name: str, order: str, expected_message: str) -> None:
    dialog.staff_name_var.set(name)
    dialog.staff_order_var.set(order)

    dialog.save_staff()

    assert messagebox.showerror.call_args.args == (constants.ERROR_TITLE, expected_message)


def test_issue_temporary_password(dialog: MasterDialog, analyzer: MedicalDocsAnalyzer,
                                  messagebox: MagicMock) -> None:
    messagebox.askyesno.return_value = True
    dialog.staff_tree.selection_set('0')
    dialog.issue_temporary_password()
    assert messagebox.showerror.call_args.args == (constants.ERROR_TITLE, constants.EMR_ID_REQUIRED_MESSAGE)

    dialog._on_staff_selected()
    dialog.staff_emr_id_var.set('A01')
    dialog.staff_role_var.set(constants.ADMIN_ROLE_LABEL)
    dialog.save_staff()
    dialog.staff_tree.selection_set('0')
    dialog.issue_temporary_password()

    password = messagebox.showinfo.call_args.args[1].split(': ')[1].split()[0]
    result = analyzer.authenticate('A01', password)
    assert result is not None and result.must_change_password
    assert analyzer.login_required()


def test_save_and_delete_department(dialog: MasterDialog, analyzer: MedicalDocsAnalyzer,
                                    messagebox: MagicMock) -> None:
    dialog.department_name_var.set('眼科')
    dialog.save_department()
    dialog.department_tree.selection_set('0')
    messagebox.askyesno.return_value = True

    dialog.delete_department()

    assert analyzer.list_departments() == [Department('外科', 2), Department('眼科', 3)]
