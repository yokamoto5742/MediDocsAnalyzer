import tkinter as tk

from app.app_window import MedicalDocsAnalyzerGUI


def main() -> None:
    root = tk.Tk()
    MedicalDocsAnalyzerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
