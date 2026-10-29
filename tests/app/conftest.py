import tkinter as tk
from collections.abc import Iterator

import pytest


@pytest.fixture(scope='session')
def session_root() -> Iterator[tk.Tk]:
    # Tk() の作成と破棄を繰り返すと Windows で Tcl の読み込みに失敗することがあるため、1 つを使い回す
    root = tk.Tk()
    root.withdraw()
    yield root
    root.destroy()


@pytest.fixture
def root(session_root: tk.Tk) -> Iterator[tk.Tk]:
    yield session_root
    for child in session_root.winfo_children():
        child.destroy()
