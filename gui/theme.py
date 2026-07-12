from tkinter import ttk


def setup_dark_theme(root) -> None:
    style = ttk.Style()
    bg_color = "#2b2b2b"
    fg_color = "#ffffff"

    style.configure("TLabel", foreground=fg_color)
    style.configure("TLabelframe", foreground=fg_color)
    style.configure("TLabelframe.Label", foreground=fg_color)
    style.configure("TButton", foreground=fg_color)
    style.configure("TRadiobutton", foreground=fg_color)
    style.configure("TCheckbutton", foreground=fg_color)

    root.configure(bg=bg_color)


DARK_BG = "#2b2b2b"
DARK_FG = "#ffffff"