import tkinter as tk
from app_window import MedicalDocsAnalyzerGUI


def main():
    root = tk.Tk()
    MedicalDocsAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
