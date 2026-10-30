import sqlite3
from collections.abc import Iterator

import pytest

from service.service_database import connect
from service.service_master import (ADMIN_ROLE, Department, Staff, delete_department, list_departments, list_staff,
                                    save_department, save_staff, seed_masters)


@pytest.fixture
def connection() -> Iterator[sqlite3.Connection]:
    connection = connect(':memory:')
    seed_masters(connection, ['植田', '花﨑'], ['内科', '外科'])
    yield connection
    connection.close()


def test_seed_masters_only_when_empty(connection: sqlite3.Connection) -> None:
    seed_masters(connection, ['別人'], ['眼科'])

    assert [(staff.name, staff.display_order) for staff in list_staff(connection)] == [('植田', 1), ('花﨑', 2)]
    assert list_departments(connection) == [Department('内科', 1), Department('外科', 2)]


def test_save_staff_adds_and_updates(connection: sqlite3.Connection) -> None:
    save_staff(connection, Staff(None, '沖野', 0, emr_id='E01', role=ADMIN_ROLE))
    okino = list_staff(connection)[0]
    okino.active = False
    okino.emr_id = ''

    save_staff(connection, okino)

    assert list_staff(connection)[0] == Staff(okino.id, '沖野', 0, active=False, emr_id='', role=ADMIN_ROLE)


def test_save_staff_allows_several_without_emr_id_but_rejects_duplicates(connection: sqlite3.Connection) -> None:
    save_staff(connection, Staff(None, '沖野', 3))
    save_staff(connection, Staff(None, '小牧', 4, emr_id='E01'))

    with pytest.raises(sqlite3.IntegrityError):
        save_staff(connection, Staff(None, '植田', 5))
    with pytest.raises(sqlite3.IntegrityError):
        save_staff(connection, Staff(None, '大代', 6, emr_id='E01'))


def test_save_and_delete_department(connection: sqlite3.Connection) -> None:
    save_department(connection, None, Department('眼科', 0))
    save_department(connection, '外科', Department('形成外科', 5))
    delete_department(connection, '内科')

    assert list_departments(connection) == [Department('眼科', 0), Department('形成外科', 5)]
