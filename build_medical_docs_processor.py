import subprocess


def build_executable():
    subprocess.run([
        "pyinstaller",
        "--name=MedicalDocsProcessor",
        "--windowed",
        "--icon=assets/MedicalDocsAnalyzer.ico",
        "--add-data", "config.ini:.",
        "service_medical_docs_processor.py"
    ])

    print(f"Executable built successfully.")


if __name__ == "__main__":
    build_executable()
