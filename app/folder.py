import os
from pathlib import Path


def open_folder_of(file_path: str | Path) -> None:
    """保存したファイルのフォルダを開く。エクスプローラーは同じフォルダの窓が開いていれば、新しく開かずにそれを前に出す"""
    os.startfile(Path(file_path).parent)
