import sqlite3
from datetime import date
from pathlib import Path

from service.service_backup import backup_database_daily


def make_database(path: Path) -> None:
    connection = sqlite3.connect(path)
    with connection:
        connection.execute('CREATE TABLE sample (value TEXT)')
        connection.execute("INSERT INTO sample VALUES ('花﨑')")
    connection.close()


def test_backup_database_daily_creates_one_backup_per_day(tmp_path: Path) -> None:
    database = tmp_path / 'database.db'
    make_database(database)
    backup_dir = tmp_path / 'backup'

    first = backup_database_daily(str(database), str(backup_dir), 30, date(2025, 4, 1))
    second = backup_database_daily(str(database), str(backup_dir), 30, date(2025, 4, 1))

    assert first == backup_dir / 'database_20250401.db'
    assert second is None
    assert first is not None
    connection = sqlite3.connect(first)
    assert connection.execute('SELECT value FROM sample').fetchall() == [('花﨑',)]
    connection.close()
    assert [path.name for path in backup_dir.iterdir()] == ['database_20250401.db']


def test_backup_database_daily_removes_old_generations(tmp_path: Path) -> None:
    database = tmp_path / 'database.db'
    make_database(database)
    backup_dir = tmp_path / 'backup'

    for day in range(1, 5):
        backup_database_daily(str(database), str(backup_dir), 2, date(2025, 4, day))

    assert sorted(path.name for path in backup_dir.iterdir()) == ['database_20250403.db', 'database_20250404.db']


def test_backup_database_daily_skips_missing_database(tmp_path: Path) -> None:
    assert backup_database_daily(str(tmp_path / 'none.db'), str(tmp_path / 'backup'), 30, date(2025, 4, 1)) is None
    assert not (tmp_path / 'backup').exists()
