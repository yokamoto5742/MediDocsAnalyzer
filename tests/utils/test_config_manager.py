import sys
from pathlib import Path

import pytest

from utils.config_manager import get_config_path, load_config, save_config


@pytest.fixture
def frozen_app(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> tuple[Path, Path]:
    bundle_dir = tmp_path / 'bundle'
    exe_dir = tmp_path / 'app'
    bundle_dir.mkdir()
    exe_dir.mkdir()
    (bundle_dir / 'config.ini').write_text('[PATHS]\n', encoding='utf-8')
    monkeypatch.setattr(sys, 'frozen', True, raising=False)
    monkeypatch.setattr(sys, '_MEIPASS', str(bundle_dir), raising=False)
    monkeypatch.setattr(sys, 'executable', str(exe_dir / 'app.exe'))
    return bundle_dir, exe_dir


def test_get_config_path_copies_bundled_ini_next_to_exe(frozen_app: tuple[Path, Path]) -> None:
    _, exe_dir = frozen_app

    assert Path(get_config_path()) == exe_dir / 'config.ini'
    assert (exe_dir / 'config.ini').read_text(encoding='utf-8') == '[PATHS]\n'


def test_get_config_path_keeps_existing_ini(frozen_app: tuple[Path, Path]) -> None:
    _, exe_dir = frozen_app
    (exe_dir / 'config.ini').write_text('[Saved]\n', encoding='utf-8')

    get_config_path()

    assert (exe_dir / 'config.ini').read_text(encoding='utf-8') == '[Saved]\n'


def test_load_config_reads_utf8(config_path: Path) -> None:
    config = load_config()

    assert config['Analysis']['ordered_names'] == '植田,沖野,花﨑'


def test_save_config_round_trip(config_path: Path) -> None:
    config = load_config()
    config['Analysis']['start_date'] = '2025-10-01'

    save_config(config)

    assert load_config()['Analysis']['start_date'] == '2025-10-01'
    assert '花﨑' in config_path.read_text(encoding='utf-8')
