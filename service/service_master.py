import sqlite3
from dataclasses import dataclass

ADMIN_ROLE = 'admin'
STAFF_ROLE = 'staff'


@dataclass
class Staff:
    id: int | None
    name: str
    display_order: int
    active: bool = True
    emr_id: str = ''
    role: str = STAFF_ROLE


@dataclass
class Department:
    name: str
    display_order: int


def seed_masters(connection: sqlite3.Connection, staff_names: list[str], departments: list[str]) -> None:
    """マスタが空のときだけ、設定ファイルの担当者・診療科を初期値として入れる"""
    if connection.execute('SELECT COUNT(*) FROM staff').fetchone()[0] == 0:
        connection.executemany('INSERT INTO staff (name, display_order) VALUES (?, ?)',
                               [(name, order) for order, name in enumerate(staff_names, start=1)])
    if connection.execute('SELECT COUNT(*) FROM departments').fetchone()[0] == 0:
        connection.executemany('INSERT INTO departments (name, display_order) VALUES (?, ?)',
                               [(name, order) for order, name in enumerate(departments, start=1)])


def list_staff(connection: sqlite3.Connection) -> list[Staff]:
    rows = connection.execute(
        'SELECT id, name, display_order, active, emr_id, role FROM staff ORDER BY display_order, id')
    return [Staff(id_, name, order, bool(active), emr_id or '', role)
            for id_, name, order, active, emr_id, role in rows]


def save_staff(connection: sqlite3.Connection, staff: Staff) -> None:
    """id が無ければ追加し、あれば更新する。氏名や電子カルテIDが重複すると sqlite3.IntegrityError"""
    # 電子カルテIDの未入力は NULL にして、UNIQUE 制約の対象から外す
    values = (staff.name, staff.display_order, int(staff.active), staff.emr_id or None, staff.role)
    if staff.id is None:
        connection.execute('INSERT INTO staff (name, display_order, active, emr_id, role) VALUES (?, ?, ?, ?, ?)',
                           values)
    else:
        connection.execute('UPDATE staff SET name = ?, display_order = ?, active = ?, emr_id = ?, role = ?'
                           ' WHERE id = ?', (*values, staff.id))


def list_departments(connection: sqlite3.Connection) -> list[Department]:
    rows = connection.execute('SELECT name, display_order FROM departments ORDER BY display_order, name')
    return [Department(name, order) for name, order in rows]


def save_department(connection: sqlite3.Connection, old_name: str | None, department: Department) -> None:
    if old_name is None:
        connection.execute('INSERT INTO departments (name, display_order) VALUES (?, ?)',
                           (department.name, department.display_order))
    else:
        connection.execute('UPDATE departments SET name = ?, display_order = ? WHERE name = ?',
                           (department.name, department.display_order, old_name))


def delete_department(connection: sqlite3.Connection, name: str) -> None:
    connection.execute('DELETE FROM departments WHERE name = ?', (name,))
