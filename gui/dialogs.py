from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from core import HistoryManager
from errors import HistoryError


def open_create_checkpoint_dialog(root, history_manager: HistoryManager, library_path_getter, on_success) -> None:
    dialog = tk.Toplevel(root)
    dialog.title("Create Checkpoint")
    dialog.geometry("400x150")
    dialog.transient(root)
    dialog.grab_set()

    ttk.Label(dialog, text="Checkpoint Label:").pack(padx=10, pady=(10, 0), anchor="w")
    label_entry = ttk.Entry(dialog, width=45)
    label_entry.pack(padx=10, pady=5, fill="x")
    label_entry.focus()

    def create() -> None:
        label = label_entry.get().strip()
        if not label:
            messagebox.showwarning("Empty Label", "Please enter a checkpoint label.")
            return

        try:
            library_dir = library_path_getter().strip()
            history_manager.add_checkpoint(label, library_dir)
            dialog.destroy()
            on_success(label)
            messagebox.showinfo("Success", f"Checkpoint '{label}' created successfully.")
        except HistoryError as exc:
            messagebox.showerror("History Error", str(exc))

    def on_enter(_event) -> None:
        create()

    button_frame = ttk.Frame(dialog)
    button_frame.pack(pady=10)
    ttk.Button(button_frame, text="Create", command=create).pack(side="left", padx=5)
    ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="left", padx=5)

    label_entry.bind("<Return>", on_enter)


def open_rename_checkpoint_dialog(root, history_manager: HistoryManager, history_index: int, current_label: str, on_success) -> None:
    dialog = tk.Toplevel(root)
    dialog.title("Rename Checkpoint")
    dialog.geometry("400x150")
    dialog.transient(root)
    dialog.grab_set()

    ttk.Label(dialog, text="New Checkpoint Label:").pack(padx=10, pady=(10, 0), anchor="w")
    label_entry = ttk.Entry(dialog, width=45)
    label_entry.insert(0, current_label)
    label_entry.pack(padx=10, pady=5, fill="x")
    label_entry.focus()
    label_entry.select_range(0, tk.END)

    def rename() -> None:
        new_label = label_entry.get().strip()
        if not new_label:
            messagebox.showwarning("Empty Label", "Please enter a checkpoint label.")
            return

        try:
            if history_manager.rename_checkpoint(history_index, new_label):
                dialog.destroy()
                on_success(new_label)
                messagebox.showinfo("Success", f"Checkpoint renamed to '{new_label}'.")
            else:
                messagebox.showerror("Error", "Failed to rename checkpoint.")
        except HistoryError as exc:
            messagebox.showerror("History Error", str(exc))

    def on_enter(_event) -> None:
        rename()

    button_frame = ttk.Frame(dialog)
    button_frame.pack(pady=10)
    ttk.Button(button_frame, text="Rename", command=rename).pack(side="left", padx=5)
    ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="left", padx=5)

    label_entry.bind("<Return>", on_enter)