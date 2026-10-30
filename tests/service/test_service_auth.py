import sqlite3
from collections.abc import Iterator

import pytest

from service.service_auth import (TEMPORARY_PASSWORD_CHARACTERS, TEMPORARY_PASSWORD_LENGTH, authenticate,
                                  login_required, make_temporary_password, set_password)
from service.service_database import connect
from service.service_master import ADMIN_ROLE, Staff, list_staff, save_staff


@pytest.fixture
def connection() -> Iterator[sqlite3.Connection]:
    connection = connect(':memory:')
    save_staff(connection, Staff(None, '植田', 1, emr_id='A01', role=ADMIN_ROLE))
    save_staff(connection, Staff(None, '花﨑', 2, emr_id='S01'))
    yield connection
    connection.close()


def staff_id(connection: sqlite3.Connection, name: str) -> int:
    staff_ids = {staff.name: staff.id for staff in list_staff(connection)}
    staff_id = staff_ids[name]
    assert staff_id is not None
    return staff_id


def test_login_required_only_when_admin_has_password(connection: sqlite3.Connection) -> None:
    assert not login_required(connection)
    set_password(connection, staff_id(connection, '花﨑'), 'secret1', must_change=False)
    assert not login_required(connection)
    set_password(connection, staff_id(connection, '植田'), 'secret1', must_change=False)
    assert login_required(connection)


def test_authenticate(connection: sqlite3.Connection) -> None:
    set_password(connection, staff_id(connection, '花﨑'), 'secret1', must_change=True)

    result = authenticate(connection, 'S01', 'secret1')

    assert result is not None
    assert (result.staff.name, result.must_change_password) == ('花﨑', True)
    assert authenticate(connection, 'S01', 'wrong') is None
    assert authenticate(connection, 'A01', 'secret1') is None


def test_password_is_not_stored_as_plain_text(connection: sqlite3.Connection) -> None:
    set_password(connection, staff_id(connection, '花﨑'), 'secret1', must_change=False)

    stored = connection.execute("SELECT password_hash FROM staff WHERE emr_id = 'S01'").fetchone()[0]

    assert 'secret1' not in stored


def test_authenticate_rejects_retired_staff(connection: sqlite3.Connection) -> None:
    hanazaki = list_staff(connection)[1]
    assert hanazaki.id is not None
    set_password(connection, hanazaki.id, 'secret1', must_change=False)
    hanazaki.active = False
    save_staff(connection, hanazaki)

    assert authenticate(connection, 'S01', 'secret1') is None


def test_make_temporary_password() -> None:
    password = make_temporary_password()

    assert len(password) == TEMPORARY_PASSWORD_LENGTH
    assert set(password) <= set(TEMPORARY_PASSWORD_CHARACTERS)
