import os
from pathlib import Path

from dotenv import load_dotenv


def load_environment_variables():
    base_dir = Path(__file__).parent.parent
    env_path = os.path.join(base_dir, '.env')

    if os.path.exists(env_path):
        load_dotenv(env_path)
        print(".envファイルを読み込みました")
    else:
        print("警告: .envファイルが見つかりません。")
