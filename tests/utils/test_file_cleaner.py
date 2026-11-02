import configparser
import os
import time
from pathlib import Path

from utils.file_cleaner import cleanup_old_files


def make_config(output_dir: Path, enabled: bool = True, retention_days: int = 1) -> configparser.ConfigParser:
    config = configparser.ConfigParser()
    config['PATHS'] = {'output_dir': str(output_dir)}
    config['FileCleanup'] = {'enabled': str(enabled).lower(), 'retention_days': str(retention_days)}
    return config


def make_file(path: Path, age_days: float) -> Path:
    path.write_text('x', encoding='utf-8')
    modified = time.time() - age_days * 24 * 60 * 60
    os.utime(path, (modified, modified))
    return path


def test_cleanup_old_files_removes_files_older_than_retention(tmp_path: Path) -> None:
    old_file = make_file(tmp_path / 'old.xlsx', 2)
    new_file = make_file(tmp_path / 'new.csv', 0.5)
    sub_dir = tmp_path / 'sub'
    sub_dir.mkdir()

    cleanup_old_files(make_config(tmp_path))

    assert not old_file.exists()
    assert new_file.exists()
    assert sub_dir.exists()


def test_cleanup_old_files_does_nothing_when_disabled(tmp_path: Path) -> None:
    old_file = make_file(tmp_path / 'old.xlsx', 2)

    cleanup_old_files(make_config(tmp_path, enabled=False))

    assert old_file.exists()


def test_cleanup_old_files_ignores_missing_output_dir(tmp_path: Path) -> None:
    cleanup_old_files(make_config(tmp_path / 'missing'))
