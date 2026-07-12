from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core.playlist_service import create_playlist
from errors.playlist_errors import PlaylistError
from gui.theme import DARK_BG, DARK_FG


class PlaylistTab:
    def __init__(self, parent, status_callback=None, progress_callback=None) -> None:
        self.parent = parent
        self.status_callback = status_callback
        self.progress_callback = progress_callback

        self.source_dir = tk.StringVar()
        self.output_path = tk.StringVar()
        self.recursive = tk.BooleanVar(value=True)

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
            text="Winamp Playlist Builder",
            bg=bg,
            fg=text,
            font=("Segoe UI", 16, "bold"),
        ).pack(anchor="w", pady=(0, 8))

        tk.Label(
            frame,
            text="Create an M3U8 playlist from audio files in a folder",
            bg=bg,
            fg=muted,
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(0, 16))

        source_frame = tk.Frame(frame, bg=panel, highlightbackground=border, highlightthickness=1)
        source_frame.pack(fill="x", pady=(0, 12))

        tk.Label(
            source_frame,
            text="Source Folder:",
            bg=panel,
            fg=text,
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", padx=12, pady=(10, 4))

        source_row = tk.Frame(source_frame, bg=panel)
        source_row.pack(fill="x", padx=12, pady=(0, 12))

        self.source_entry = tk.Entry(
            source_row,
            textvariable=self.source_dir,
            bg="#1b1b1c",
            fg=text,
            insertbackground=text,
            relief="flat",
            highlightthickness=1,
            highlightbackground=border,
            highlightcolor=accent,
            font=("Consolas", 10),
        )
        self.source_entry.pack(side="left", fill="x", expand=True, ipady=8)

        self.source_browse_btn = tk.Button(
            source_row,
            text="Browse",
            command=self._browse_source,
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
        self.source_browse_btn.pack(side="left", padx=(10, 0))

        output_frame = tk.Frame(frame, bg=panel, highlightbackground=border, highlightthickness=1)
        output_frame.pack(fill="x", pady=(0, 12))

        tk.Label(
            output_frame,
            text="Playlist File (.m3u8):",
            bg=panel,
            fg=text,
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", padx=12, pady=(10, 4))

        output_row = tk.Frame(output_frame, bg=panel)
        output_row.pack(fill="x", padx=12, pady=(0, 12))

        self.output_entry = tk.Entry(
            output_row,
            textvariable=self.output_path,
            bg="#1b1b1c",
            fg=text,
            insertbackground=text,
            relief="flat",
            highlightthickness=1,
            highlightbackground=border,
            highlightcolor=accent,
            font=("Consolas", 10),
        )
        self.output_entry.pack(side="left", fill="x", expand=True, ipady=8)

        self.output_browse_btn = tk.Button(
            output_row,
            text="Save As",
            command=self._browse_output,
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
        self.output_browse_btn.pack(side="left", padx=(10, 0))

        options_frame = tk.Frame(frame, bg=panel, highlightbackground=border, highlightthickness=1)
        options_frame.pack(fill="x", pady=(0, 12))

        tk.Label(
            options_frame,
            text="Options:",
            bg=panel,
            fg=text,
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", padx=12, pady=(10, 4))

        tk.Checkbutton(
            options_frame,
            text="Include subfolders",
            variable=self.recursive,
            bg=panel,
            fg=text,
            activebackground=panel,
            activeforeground=text,
            selectcolor="#1b1b1c",
            font=("Segoe UI", 10),
        ).pack(anchor="w", padx=12, pady=(0, 12))

        btn_row = tk.Frame(frame, bg=bg)
        btn_row.pack(fill="x", pady=(0, 12))

        self.start_btn = tk.Button(
            btn_row,
            text="Create Playlist",
            command=self._create_playlist,
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

        self.write_log("Playlist tab ready.")
        self.write_log("Select a source folder and create an M3U8 playlist.")

    def write_log(self, text: str) -> None:
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.frame.update_idletasks()

    def set_local_status(self, text: str, color=None) -> None:
        self.status_label.config(text=text, fg=color or self.colors["muted"])
        self._set_status(text)
        self.frame.update_idletasks()

    def _browse_source(self) -> None:
        folder = filedialog.askdirectory(
            parent=self.frame,
            mustexist=True,
            title="Select source folder",
        )
        if folder:
            self.source_dir.set(folder)
            if not self.output_path.get().strip():
                folder_name = folder.rstrip("\\/").split("/")[-1].split("\\")[-1]
                self.output_path.set(f"{folder}/{folder_name}.m3u8")

    def _browse_output(self) -> None:
        file_path = filedialog.asksaveasfilename(
            parent=self.frame,
            title="Save playlist as",
            defaultextension=".m3u8",
            filetypes=[("M3U playlist", "*.m3u"), ("All files", "*.*")],
        )
        if file_path:
            self.output_path.set(file_path)

    def _clear_log(self) -> None:
        self.log.delete("1.0", tk.END)
        self.write_log("Log cleared.")

    def _create_playlist(self) -> None:
        source_folder = self.source_dir.get().strip()
        output_path = self.output_path.get().strip()
        recursive = self.recursive.get()

        if not source_folder:
            messagebox.showerror("Playlist Builder", "Choose a source folder first.")
            return

        self.log.delete("1.0", tk.END)
        self.set_local_status("Creating playlist...", self.colors["accent"])

        try:
            playlist_path, item_count = create_playlist(
                source_folder=source_folder,
                output_path=output_path or None,
                recursive=recursive,
                progress_callback=self._set_progress,
                log_callback=self.write_log,
            )

            self._set_progress(100, 100, "Done.")
            self.set_local_status(f"Done. Files added: {item_count}", self.colors["ok"])
            self.write_log("")
            self.write_log(f"Playlist created: {playlist_path}")
            self.write_log(f"Files added: {item_count}")

            messagebox.showinfo(
                "Playlist Created",
                f"Playlist saved successfully.\n\nPath: {playlist_path}\nFiles added: {item_count}",
            )

        except PlaylistError as exc:
            self.write_log(f"FAIL  | {exc}")
            self.set_local_status(str(exc), self.colors["err"])
            messagebox.showerror("Playlist Error", str(exc))
        except Exception as exc:
            self.write_log(f"FAIL  | {exc}")
            self.set_local_status("Unexpected error.", self.colors["err"])
            messagebox.showerror("Unexpected Error", str(exc))