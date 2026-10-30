import sqlite3
import tkinter as tk
from tkinter import messagebox, ttk

from app import constants
from service.service_master import ADMIN_ROLE, STAFF_ROLE, Department, Staff
from service.service_medical_docs_analyzer import MedicalDocsAnalyzer

ROLE_LABELS = {ADMIN_ROLE: constants.ADMIN_ROLE_LABEL, STAFF_ROLE: constants.STAFF_ROLE_LABEL}
ROLES_BY_LABEL = {label: role for role, label in ROLE_LABELS.items()}


def make_tree(parent: ttk.Frame, columns: tuple[str, ...], widths: tuple[int, ...]) -> ttk.Treeview:
    tree = ttk.Treeview(parent, columns=columns, show='headings', height=15, selectmode='browse')
    for column, width in zip(columns, widths):
        tree.heading(column, text=column)
        tree.column(column, width=width)
    tree.grid(row=0, column=0, columnspan=4, sticky=tk.NSEW, pady=5)
    return tree


def add_field(parent: ttk.Frame, row: int, label: str, widget: tk.Widget) -> None:
    ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, padx=5, pady=2)
    widget.grid(row=row, column=1, columnspan=3, sticky=tk.W, padx=5, pady=2)


def parse_order(text: str) -> int:
    try:
        return int(text)
    except ValueError:
        raise ValueError(constants.INVALID_ORDER_MESSAGE) from None


class MasterDialog:
    """担当者と診療科のマスタを編集する"""

    def __init__(self, parent: tk.Tk, analyzer: MedicalDocsAnalyzer) -> None:
        self.analyzer = analyzer
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(constants.MASTER_TITLE)
        self.dialog.transient(parent)
        self.dialog.grab_set()

        notebook = ttk.Notebook(self.dialog, padding="10")
        notebook.grid(row=0, column=0)
        self._setup_staff_tab(notebook)
        self._setup_department_tab(notebook)
        ttk.Button(self.dialog, text=constants.CLOSE_BUTTON_LABEL, command=self.dialog.destroy).grid(
            row=1, column=0, pady=5)

        self.staff_list: list[Staff] = []
        self.departments: list[Department] = []
        self.refresh()

    def _setup_staff_tab(self, notebook: ttk.Notebook) -> None:
        frame = ttk.Frame(notebook, padding="5")
        notebook.add(frame, text=constants.STAFF_TAB_LABEL)
        self.staff_tree = make_tree(frame, constants.STAFF_COLUMNS, (120, 60, 60, 120, 80))
        self.staff_tree.bind('<<TreeviewSelect>>', self._on_staff_selected)

        self.staff_name_var = tk.StringVar()
        self.staff_order_var = tk.StringVar()
        self.staff_active_var = tk.BooleanVar(value=True)
        self.staff_emr_id_var = tk.StringVar()
        self.staff_role_var = tk.StringVar(value=constants.STAFF_ROLE_LABEL)
        add_field(frame, 1, constants.STAFF_COLUMNS[0], ttk.Entry(frame, textvariable=self.staff_name_var))
        add_field(frame, 2, constants.STAFF_COLUMNS[1], ttk.Entry(frame, width=6, textvariable=self.staff_order_var))
        add_field(frame, 3, constants.STAFF_COLUMNS[2], ttk.Checkbutton(frame, variable=self.staff_active_var))
        add_field(frame, 4, constants.STAFF_COLUMNS[3], ttk.Entry(frame, textvariable=self.staff_emr_id_var))
        add_field(frame, 5, constants.STAFF_COLUMNS[4], ttk.Combobox(
            frame, width=8, state='readonly', values=list(ROLES_BY_LABEL), textvariable=self.staff_role_var))

        self.staff_button_frame = ttk.Frame(frame)
        self.staff_button_frame.grid(row=6, column=0, columnspan=4, pady=5)
        ttk.Button(self.staff_button_frame, text=constants.NEW_BUTTON_LABEL, command=self.clear_staff_form).grid(
            row=0, column=0, padx=5)
        ttk.Button(self.staff_button_frame, text=constants.SAVE_BUTTON_LABEL, command=self.save_staff).grid(
            row=0, column=1, padx=5)
        ttk.Button(self.staff_button_frame, text=constants.ISSUE_PASSWORD_BUTTON_LABEL,
                   command=self.issue_temporary_password).grid(row=0, column=2, padx=5)

    def _setup_department_tab(self, notebook: ttk.Notebook) -> None:
        frame = ttk.Frame(notebook, padding="5")
        notebook.add(frame, text=constants.DEPARTMENT_TAB_LABEL)
        self.department_tree = make_tree(frame, constants.DEPARTMENT_COLUMNS, (160, 60))
        self.department_tree.bind('<<TreeviewSelect>>', self._on_department_selected)

        self.department_name_var = tk.StringVar()
        self.department_order_var = tk.StringVar()
        add_field(frame, 1, constants.DEPARTMENT_COLUMNS[0], ttk.Entry(frame, textvariable=self.department_name_var))
        add_field(frame, 2, constants.DEPARTMENT_COLUMNS[1],
                  ttk.Entry(frame, width=6, textvariable=self.department_order_var))

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=3, column=0, columnspan=4, pady=5)
        for column, (label, command) in enumerate((
                (constants.NEW_BUTTON_LABEL, self.clear_department_form),
                (constants.SAVE_BUTTON_LABEL, self.save_department),
                (constants.DELETE_BUTTON_LABEL, self.delete_department))):
            ttk.Button(button_frame, text=label, command=command).grid(row=0, column=column, padx=5)

    def refresh(self) -> None:
        self.staff_list = self.analyzer.list_staff()
        self.departments = self.analyzer.list_departments()
        self.staff_tree.delete(*self.staff_tree.get_children())
        for index, staff in enumerate(self.staff_list):
            self.staff_tree.insert('', tk.END, iid=str(index), values=(
                staff.name, staff.display_order, constants.ACTIVE_LABEL if staff.active else constants.RETIRED_LABEL,
                staff.emr_id, ROLE_LABELS[staff.role]))
        self.department_tree.delete(*self.department_tree.get_children())
        for index, department in enumerate(self.departments):
            self.department_tree.insert('', tk.END, iid=str(index),
                                        values=(department.name, department.display_order))
        self.clear_staff_form()
        self.clear_department_form()

    def selected_staff(self) -> Staff | None:
        selection = self.staff_tree.selection()
        return self.staff_list[int(selection[0])] if selection else None

    def selected_department(self) -> Department | None:
        selection = self.department_tree.selection()
        return self.departments[int(selection[0])] if selection else None

    def _on_staff_selected(self, _event: tk.Event | None = None) -> None:
        staff = self.selected_staff()
        if staff is None:
            return
        self.staff_name_var.set(staff.name)
        self.staff_order_var.set(str(staff.display_order))
        self.staff_active_var.set(staff.active)
        self.staff_emr_id_var.set(staff.emr_id)
        self.staff_role_var.set(ROLE_LABELS[staff.role])

    def _on_department_selected(self, _event: tk.Event | None = None) -> None:
        department = self.selected_department()
        if department is None:
            return
        self.department_name_var.set(department.name)
        self.department_order_var.set(str(department.display_order))

    def clear_staff_form(self) -> None:
        self.staff_tree.selection_remove(*self.staff_tree.selection())
        self.staff_name_var.set('')
        next_order = max((staff.display_order for staff in self.staff_list), default=0) + 1
        self.staff_order_var.set(str(next_order))
        self.staff_active_var.set(True)
        self.staff_emr_id_var.set('')
        self.staff_role_var.set(constants.STAFF_ROLE_LABEL)

    def clear_department_form(self) -> None:
        self.department_tree.selection_remove(*self.department_tree.selection())
        self.department_name_var.set('')
        next_order = max((department.display_order for department in self.departments), default=0) + 1
        self.department_order_var.set(str(next_order))

    def save_staff(self) -> None:
        selected = self.selected_staff()
        name = self.staff_name_var.get().strip()
        if not name:
            messagebox.showerror(constants.ERROR_TITLE, constants.NAME_REQUIRED_MESSAGE, parent=self.dialog)
            return
        try:
            staff = Staff(selected.id if selected else None, name, parse_order(self.staff_order_var.get()),
                          self.staff_active_var.get(), self.staff_emr_id_var.get().strip(),
                          ROLES_BY_LABEL[self.staff_role_var.get()])
            self.analyzer.save_staff(staff)
        except ValueError as e:
            messagebox.showerror(constants.ERROR_TITLE, str(e), parent=self.dialog)
            return
        except sqlite3.IntegrityError:
            messagebox.showerror(constants.ERROR_TITLE, constants.DUPLICATE_STAFF_MESSAGE, parent=self.dialog)
            return
        self.refresh()

    def issue_temporary_password(self) -> None:
        staff = self.selected_staff()
        if staff is None or staff.id is None:
            return
        if not staff.emr_id:
            messagebox.showerror(constants.ERROR_TITLE, constants.EMR_ID_REQUIRED_MESSAGE, parent=self.dialog)
            return
        if not messagebox.askyesno(constants.CONFIRM_TITLE,
                                   constants.ISSUE_PASSWORD_CONFIRM_MESSAGE.format(name=staff.name),
                                   parent=self.dialog):
            return
        password = self.analyzer.issue_temporary_password(staff.id)
        messagebox.showinfo(constants.ISSUE_PASSWORD_BUTTON_LABEL,
                            constants.TEMPORARY_PASSWORD_MESSAGE.format(name=staff.name, password=password),
                            parent=self.dialog)

    def save_department(self) -> None:
        selected = self.selected_department()
        name = self.department_name_var.get().strip()
        if not name:
            messagebox.showerror(constants.ERROR_TITLE, constants.NAME_REQUIRED_MESSAGE, parent=self.dialog)
            return
        try:
            department = Department(name, parse_order(self.department_order_var.get()))
            self.analyzer.save_department(selected.name if selected else None, department)
        except ValueError as e:
            messagebox.showerror(constants.ERROR_TITLE, str(e), parent=self.dialog)
            return
        except sqlite3.IntegrityError:
            messagebox.showerror(constants.ERROR_TITLE, constants.DUPLICATE_DEPARTMENT_MESSAGE, parent=self.dialog)
            return
        self.refresh()

    def delete_department(self) -> None:
        department = self.selected_department()
        if department is None:
            return
        if messagebox.askyesno(constants.CONFIRM_TITLE,
                               constants.DELETE_DEPARTMENT_CONFIRM_MESSAGE.format(name=department.name),
                               parent=self.dialog):
            self.analyzer.delete_department(department.name)
            self.refresh()
