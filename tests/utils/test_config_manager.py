from pathlib import Path

from utils.config_manager import load_config, save_config


def test_load_config_reads_utf8(config_path: Path) -> None:
    config = load_config()

    assert config['Analysis']['ordered_names'] == '植田,沖野,花﨑'


def test_save_config_round_trip(config_path: Path) -> None:
    config = load_config()
    config['Analysis']['start_date'] = '2025-10-01'

    save_config(config)

    assert load_config()['Analysis']['start_date'] == '2025-10-01'
    assert '花﨑' in config_path.read_text(encoding='utf-8')
