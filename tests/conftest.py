from pathlib import Path

import pytest

from utils import config_manager


@pytest.fixture
def config_path(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """一時フォルダの設定ファイルを使う(本物の config.ini を書き換えない)"""
    path = tmp_path / 'config.ini'
    path.write_text(f"""[Appearance]
window_width = 300
window_height = 450

[PATHS]
source_file_path = {tmp_path / 'source.xlsx'}
database_path = {tmp_path / 'database.db'}
legacy_database_path = {tmp_path / 'legacy.xlsx'}
backup_dir = {tmp_path / 'backup'}
template_path = {tmp_path / 'template.xlsx'}
output_dir = {tmp_path / 'output'}

[Backup]
keep_generations = 30

[Analysis]
start_date = 2025-04-01
end_date = 2025-06-30
year_start_month = 4
ordered_names = 植田,沖野,花﨑
clinical_departments = 合計,内科,外科
""", encoding='utf-8')
    monkeypatch.setattr(config_manager, 'CONFIG_PATH', str(path))
    return path
