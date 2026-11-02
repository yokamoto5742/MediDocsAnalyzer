from pathlib import Path
from unittest.mock import patch

from app.folder import open_folder_of


def test_open_folder_of_opens_parent_folder(tmp_path: Path) -> None:
    with patch('app.folder.os.startfile', create=True) as startfile:
        open_folder_of(str(tmp_path / 'output' / 'result.xlsx'))

    startfile.assert_called_once_with(tmp_path / 'output')
