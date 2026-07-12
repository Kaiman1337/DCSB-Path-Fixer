import tkinter as tk

from gui.app import DCSBPathFixerApp


def run_app() -> None:
    root = tk.Tk()
    root.option_add("*tearOff", False)
    DCSBPathFixerApp(root)
    root.mainloop()