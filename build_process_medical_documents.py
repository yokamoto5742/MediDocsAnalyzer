import subprocess
from version_manager import update_version, update_version_py


def build_executable():
    new_version = update_version()
    update_version_py(new_version)

    subprocess.run([
        "pyinstaller",
        "--name=MedicalDocsAnalyzer",
        "--windowed",
        "--icon=assets/MedicalDocsAnalyzer.ico",
        "--add-data", "config.ini:.",
        "process_medical_documents.py"
    ])

    print(f"Executable built successfully. Version: {new_version}")


if __name__ == "__main__":
    build_executable()
