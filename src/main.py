import tkinter as tk

from app.app_window import TaskAnalyzerGUI


def main() -> None:
    root = tk.Tk()
    TaskAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()