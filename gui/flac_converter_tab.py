from __future__ import annotations

import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core.flac_converter_service import convert_directory, ensure_dependencies
from errors.converter_errors import ConverterDependencyError, ConverterError
from gui.theme import DARK_BG, DARK_FG


class FlacConverterTab:
    def __init__(self, parent, status_callback=None, progress_callback=None) -> None:
        self.parent = parent
        self.status_callback = status_callback
        self.progress_callback = progress_callback

        self.running = False
        self.selected_dir = tk.StringVar()

        self.frame = ttk.Frame(parent)
        self._build_ui()

    def _set_status(self, text: str) -> None:
        if self.status_callback is not None:
            self.status_callback(text)

    def _set_progress(self, current: int, total: int, message: str) -> None:
        if self.progress_callback is not None:
            self.progress_callback(current, total, message)

    def _build_ui(self) -> None:
        bg = DARK_BG
        panel = "#252526"
        text = DARK_FG
        muted = "#b8b8b8"
        accent = "#3a86ff"
        btn = "#2d2d30"
        btn_hover = "#3a3a3f"
        border = "#3c3c3c"

        self.colors = {
            "bg": bg,
            "panel": panel,
            "text": text,
            "muted": muted,
            "accent": accent,
            "btn": btn,
            "btn_hover": btn_hover,
            "border": border,
            "ok": "#4caf50",
            "err": "#ff6b6b",
        }

        frame = tk.Frame(self.frame, bg=bg)
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        tk.Label(
            frame,
            text="FLAC -> MP3 Converter for Winamp",
            bg=bg,
            fg=text,
            font=("Segoe UI", 16, "bold"),
        ).pack(anchor="w", pady=(0, 8))

        tk.Label(
            frame,
            text="CBR 320k, ID3v2.3, metadata kept, verify before delete",
            bg=bg,
            fg=muted,
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(0, 16))

        path_frame = tk.Frame(frame, bg=panel, highlightbackground=border, highlightthickness=1)
        path_frame.pack(fill="x", pady=(0, 12))

        tk.Label(
            path_frame,
            text="Folder:",
            bg=panel,
            fg=text,
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", padx=12, pady=(10, 4))

        path_row = tk.Frame(path_frame, bg=panel)
        path_row.pack(fill="x", padx=12, pady=(0, 12))

        self.path_entry = tk.Entry(
            path_row,
            textvariable=self.selected_dir,
            bg="#1b1b1c",
            fg=text,
            insertbackground=text,
            relief="flat",
            highlightthickness=1,
            highlightbackground=border,
            highlightcolor=accent,
            font=("Consolas", 10),
        )
        self.path_entry.pack(side="left", fill="x", expand=True, ipady=8)

        self.browse_btn = tk.Button(
            path_row,
            text="Browse",
            command=self._browse_folder,
            bg=btn,
            fg=text,
            activebackground=btn_hover,
            activeforeground=text,
            relief="flat",
            bd=0,
            font=("Segoe UI", 10),
            padx=16,
            pady=8,
        )
        self.browse_btn.pack(side="left", padx=(10, 0))

        btn_row = tk.Frame(frame, bg=bg)
        btn_row.pack(fill="x", pady=(0, 12))

        self.start_btn = tk.Button(
            btn_row,
            text="Start",
            command=self._start_conversion,
            bg=accent,
            fg="white",
            activebackground="#2f6fd1",
            activeforeground="white",
            relief="flat",
            bd=0,
            font=("Segoe UI", 10, "bold"),
            padx=20,
            pady=10,
        )
        self.start_btn.pack(side="left")

        self.clear_btn = tk.Button(
            btn_row,
            text="Clear Log",
            command=self._clear_log,
            bg=btn,
            fg=text,
            activebackground=btn_hover,
            activeforeground=text,
            relief="flat",
            bd=0,
            font=("Segoe UI", 10),
            padx=20,
            pady=10,
        )
        self.clear_btn.pack(side="left", padx=(10, 0))

        self.status_label = tk.Label(
            frame,
            text="Ready.",
            bg=bg,
            fg=muted,
            font=("Segoe UI", 10),
        )
        self.status_label.pack(anchor="w", pady=(0, 8))

        log_frame = tk.Frame(frame, bg=panel, highlightbackground=border, highlightthickness=1)
        log_frame.pack(fill="both", expand=True)

        self.log = tk.Text(
            log_frame,
            bg="#111112",
            fg=text,
            insertbackground=text,
            relief="flat",
            bd=0,
            wrap="word",
            font=("Consolas", 10),
            padx=10,
            pady=10,
        )
        self.log.pack(side="left", fill="both", expand=True)

        scroll = tk.Scrollbar(log_frame, command=self.log.yview)
        scroll.pack(side="right", fill="y")
        self.log.config(yscrollcommand=scroll.set)

        self.write_log("Converter tab ready.")
        self.write_log("Select a folder and press Start.")

    def write_log(self, text: str) -> None:
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.frame.update_idletasks()

    def set_local_status(self, text: str, color=None) -> None:
        self.status_label.config(text=text, fg=color or self.colors["muted"])
        self._set_status(text)
        self.frame.update_idletasks()

    def _browse_folder(self) -> None:
        folder = filedialog.askdirectory(
            parent=self.frame,
            mustexist=True,
            title="Select folder with FLAC files",
        )
        if folder:
            self.selected_dir.set(folder)

    def _set_controls_state(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        self.path_entry.config(state=state)
        self.browse_btn.config(state=state)
        self.start_btn.config(state=state)

    def _clear_log(self) -> None:
        self.log.delete("1.0", tk.END)
        self.write_log("Log cleared.")

    def _start_conversion(self) -> None:
        if self.running:
            return

        folder = self.selected_dir.get().strip()
        if not folder:
            messagebox.showerror("FLAC -> MP3 Converter", "Choose a folder first.")
            return

        try:
            ensure_dependencies()
        except ConverterDependencyError as exc:
            messagebox.showerror("Dependency Error", str(exc))
            return

        worker = threading.Thread(target=self._run_conversion, args=(folder,), daemon=True)
        self.running = True
        self._set_controls_state(False)
        worker.start()

    def _run_conversion(self, folder: str) -> None:
        try:
            self.write_log(f"Folder: {folder}")
            self.set_local_status("Conversion started...", self.colors["accent"])

            ok_count, fail_count = convert_directory(
                folder,
                progress_callback=self._set_progress,
                log_callback=self.write_log,
            )

            final_color = self.colors["ok"] if fail_count == 0 else self.colors["err"]
            self.set_local_status(f"Done. OK: {ok_count}, Failed: {fail_count}", final_color)
            self.write_log("")
            self.write_log(f"Finished. OK: {ok_count}, Failed: {fail_count}")

        except ConverterError as exc:
            self.write_log(f"FAIL  | {exc}")
            self.set_local_status(str(exc), self.colors["err"])
        except Exception as exc:
            self.write_log(f"FAIL  | {exc}")
            self.set_local_status("Unexpected error.", self.colors["err"])
        finally:
            self.running = False
            self._set_controls_state(True)