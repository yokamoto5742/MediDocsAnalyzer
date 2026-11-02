import os
import time
from pathlib import Path

from utils.log_rotation import LOG_FILE_NAME, cleanup_old_logs


def make_log(path: Path, age_days: float) -> Path:
    path.write_text('x', encoding='utf-8')
    modified = time.time() - age_days * 24 * 60 * 60
    os.utime(path, (modified, modified))
    return path


def test_cleanup_old_logs_keeps_current_log_and_recent_logs(tmp_path: Path) -> None:
    current = make_log(tmp_path / LOG_FILE_NAME, 30)
    old = make_log(tmp_path / f'{LOG_FILE_NAME}.2025-01-01.log', 8)
    recent = make_log(tmp_path / f'{LOG_FILE_NAME}.2025-01-07.log', 6)
    other = make_log(tmp_path / 'memo.txt', 30)

    cleanup_old_logs(str(tmp_path), 7)

    assert current.exists()
    assert not old.exists()
    assert recent.exists()
    assert other.exists()
