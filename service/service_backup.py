import os
import sqlite3
from datetime import date
from pathlib import Path

from service.service_database import BUSY_TIMEOUT_SECONDS


def backup_database_daily(database_path: str, backup_dir: str, keep_generations: int, today: date) -> Path | None:
    """その日のバックアップがまだ無ければ作り、古いものは keep_generations 個を残して消す。

    作ったバックアップのパスを返す。DB が無いか、その日の分が作成済みなら None を返す。
    """
    database = Path(database_path)
    if not database.exists():
        return None

    backup_dir_path = Path(backup_dir)
    backup_dir_path.mkdir(parents=True, exist_ok=True)
    backup_path = backup_dir_path / f'{database.stem}_{today:%Y%m%d}{database.suffix}'
    if backup_path.exists():
        return None

    # 別の端末が同時に起動しても壊れたファイルを残さないよう、一時ファイルに書いてから名前を変える
    temp_path = backup_path.with_name(f'{backup_path.name}.{os.getpid()}.tmp')
    source = sqlite3.connect(database_path, timeout=BUSY_TIMEOUT_SECONDS)
    target = sqlite3.connect(temp_path)
    try:
        source.backup(target)
    finally:
        target.close()
        source.close()
    os.replace(temp_path, backup_path)

    backups = sorted(backup_dir_path.glob(f'{database.stem}_' + '[0-9]' * 8 + database.suffix))
    for old_backup in backups[:-keep_generations]:
        old_backup.unlink()
    return backup_path
