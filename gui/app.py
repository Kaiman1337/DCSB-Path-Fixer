from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core import DCSBPathFixerService, HistoryManager, SettingsManager
from errors import ConfigReadError, ConfigWriteError, HistoryError, SettingsError, ValidationError
from gui.constants import SUPPORTED_FORMATS_TEXT, VALID_RENAME_MODES
from gui.dialogs import open_create_checkpoint_dialog, open_rename_checkpoint_dialog
from gui.theme import DARK_BG, DARK_FG, setup_dark_theme

from gui.flac_converter_tab import FlacConverterTab
from gui.playlist_tab import PlaylistTab


class DCSBPathFixerApp:
    VIEW_SIZES = {
        "path_fixer": "1040x760",
        "playlist": "900x700",
        "flac_converter": "900x620",
    }

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("DCSB Config Path Fixer")
        self.root.geometry(self.VIEW_SIZES["path_fixer"])
        self.root.resizable(False, False)

        setup_dark_theme(self.root)

        self.service = DCSBPathFixerService()
        self.settings = SettingsManager()
        self.history_manager = HistoryManager()

        self.library_path = tk.StringVar()
        self.config_path = tk.StringVar()
        self.rename_mode = tk.StringVar(value="none")
        self.use_lowercase = tk.BooleanVar(value=False)
        self.status_text = tk.StringVar(value="Ready.")
        self.progress_value = tk.DoubleVar(value=0.0)
        self.selected_history_index = tk.IntVar(value=-1)
        self.use_config = tk.BooleanVar(value=True)
        self.active_view = tk.StringVar(value="path_fixer")

        self.log_widget: tk.Text
        self.history_listbox: tk.Listbox
        self.history_details_widget: tk.Text
        self.missing_widget: tk.Text

        self.path_fixer_view: ttk.Frame
        self.playlist_host_view: ttk.Frame
        self.flac_host_view: ttk.Frame

        self.playlist_tab: PlaylistTab | None = None
        self.flac_converter_tab: FlacConverterTab | None = None

        self._build_ui()
        self._load_settings_into_ui()
        self._bind_events()
        self._show_view("path_fixer")

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        top_tabs_frame = ttk.Frame(self.root, padding=(12, 10, 12, 0))
        top_tabs_frame.grid(row=0, column=0, sticky="ew")
        top_tabs_frame.columnconfigure(0, weight=1)
        top_tabs_frame.columnconfigure(1, weight=1)
        top_tabs_frame.columnconfigure(2, weight=1)

        self.path_fixer_tab_button = ttk.Button(
            top_tabs_frame,
            text="Path Fixer",
            command=lambda: self._show_view("path_fixer"),
        )
        self.path_fixer_tab_button.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        self.playlist_tab_button = ttk.Button(
            top_tabs_frame,
            text="Playlist Builder",
            command=lambda: self._show_view("playlist"),
        )
        self.playlist_tab_button.grid(row=0, column=1, sticky="ew", padx=4)

        self.flac_tab_button = ttk.Button(
            top_tabs_frame,
            text="FLAC -> MP3",
            command=lambda: self._show_view("flac_converter"),
        )
        self.flac_tab_button.grid(row=0, column=2, sticky="ew", padx=(4, 0))

        view_host = ttk.Frame(self.root)
        view_host.grid(row=1, column=0, sticky="nsew")
        view_host.columnconfigure(0, weight=1)
        view_host.rowconfigure(0, weight=1)

        self.path_fixer_view = ttk.Frame(view_host)
        self.playlist_host_view = ttk.Frame(view_host)
        self.flac_host_view = ttk.Frame(view_host)

        for frame in (self.path_fixer_view, self.playlist_host_view, self.flac_host_view):
            frame.grid(row=0, column=0, sticky="nsew")
            frame.grid_remove()

        self._build_path_fixer_view()
        self._build_playlist_view()
        self._build_flac_view()

    def _build_path_fixer_view(self) -> None:
        container = ttk.Frame(self.path_fixer_view, padding=12)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="Audio library folder:").grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(container, textvariable=self.library_path, width=100).grid(
            row=0, column=1, sticky="ew", padx=8, pady=(0, 6)
        )
        ttk.Button(container, text="Browse...", command=self._browse_library).grid(row=0, column=2, pady=(0, 6))

        ttk.Checkbutton(
            container,
            text="Update DCSB config.xml",
            variable=self.use_config,
            command=self._toggle_config_field,
        ).grid(row=1, column=0, sticky="w", pady=(0, 6))

        self.config_entry = ttk.Entry(container, textvariable=self.config_path, width=100)
        self.config_entry.grid(row=1, column=1, sticky="ew", padx=8, pady=(0, 6))

        self.config_browse_btn = ttk.Button(container, text="Browse...", command=self._browse_config)
        self.config_browse_btn.grid(row=1, column=2, pady=(0, 6))

        info_frame = ttk.LabelFrame(container, text="Supported audio formats", padding=8)
        info_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 10))

        info_text = tk.Text(
            info_frame,
            height=2,
            width=100,
            wrap="word",
            bg=DARK_BG,
            fg=DARK_FG,
            insertbackground=DARK_FG,
        )
        info_text.insert("1.0", SUPPORTED_FORMATS_TEXT)
        info_text.config(state="disabled")
        info_text.pack(fill="both", expand=True)

        rename_frame = ttk.LabelFrame(container, text="Rename options", padding=10)
        rename_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        rename_options = [
            ("Do not rename files", "none"),
            ("Normalize names and use spaces", "space"),
            ("Normalize names and use hyphen (-)", "-"),
            ("Normalize names and use underscore (_)", "_"),
        ]

        for text, value in rename_options:
            ttk.Radiobutton(rename_frame, text=text, variable=self.rename_mode, value=value).pack(anchor="w")

        ttk.Checkbutton(
            rename_frame, text="Convert filenames to lowercase", variable=self.use_lowercase
        ).pack(anchor="w", pady=(8, 0))

        button_frame = ttk.Frame(container)
        button_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        ttk.Button(button_frame, text="Run", command=self._run).pack(side="left")
        ttk.Button(button_frame, text="Clear output", command=self._clear_output).pack(side="left", padx=8)

        notebook = ttk.Notebook(container)
        notebook.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(4, 10))

        log_frame = ttk.Frame(notebook)
        notebook.add(log_frame, text="Log Output")

        self.log_widget = tk.Text(
            log_frame,
            height=16,
            wrap="word",
            bg=DARK_BG,
            fg=DARK_FG,
            insertbackground=DARK_FG,
        )
        self.log_widget.pack(fill="both", expand=True)

        history_frame = ttk.Frame(notebook)
        notebook.add(history_frame, text="Rename History")

        history_list_frame = ttk.LabelFrame(history_frame, text="Rename Operations", padding=4)
        history_list_frame.pack(fill="both", expand=True, side="left", padx=4, pady=4)

        scrollbar = ttk.Scrollbar(history_list_frame)
        scrollbar.pack(side="right", fill="y")

        self.history_listbox = tk.Listbox(
            history_list_frame,
            yscrollcommand=scrollbar.set,
            height=16,
            bg=DARK_BG,
            fg=DARK_FG,
            selectbackground="#505050",
            selectforeground=DARK_FG,
        )
        self.history_listbox.pack(fill="both", expand=True)
        scrollbar.config(command=self.history_listbox.yview)
        self.history_listbox.bind("<<ListboxSelect>>", self._on_history_select)

        history_detail_frame = ttk.LabelFrame(history_frame, text="Operation Details", padding=4)
        history_detail_frame.pack(fill="both", expand=True, side="right", padx=4, pady=4)

        self.history_details_widget = tk.Text(
            history_detail_frame,
            height=16,
            wrap="word",
            width=40,
            bg=DARK_BG,
            fg=DARK_FG,
            insertbackground=DARK_FG,
        )
        self.history_details_widget.pack(fill="both", expand=True)

        revert_button_frame = ttk.Frame(history_frame)
        revert_button_frame.pack(fill="x", padx=4, pady=4)
        ttk.Button(revert_button_frame, text="Create Checkpoint", command=self._create_checkpoint).pack(side="top", fill="x", pady=2)
        ttk.Button(revert_button_frame, text="Rename Checkpoint", command=self._rename_checkpoint).pack(side="top", fill="x", pady=2)
        ttk.Button(revert_button_frame, text="Revert to Selected Point", command=self._revert_to_history).pack(side="top", fill="x", pady=2)
        ttk.Button(revert_button_frame, text="Delete Selected", command=self._delete_history_entry).pack(side="top", fill="x", pady=2)
        ttk.Button(revert_button_frame, text="Refresh History", command=self._refresh_history).pack(side="top", fill="x", pady=2)

        self._refresh_history()

        ttk.Label(container, text="Files not found:").grid(row=6, column=0, sticky="w", pady=(4, 0))
        self.missing_widget = tk.Text(
            container,
            height=8,
            wrap="word",
            bg=DARK_BG,
            fg=DARK_FG,
            insertbackground=DARK_FG,
        )
        self.missing_widget.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=(4, 8))

        status_frame = ttk.Frame(container)
        status_frame.grid(row=8, column=0, columnspan=3, sticky="ew")
        status_frame.columnconfigure(1, weight=1)

        ttk.Label(status_frame, textvariable=self.status_text, relief="sunken", anchor="w").grid(
            row=0, column=0, sticky="ew", padx=(0, 8)
        )

        progress_bar = ttk.Progressbar(
            status_frame,
            variable=self.progress_value,
            maximum=100,
            mode="determinate",
        )
        progress_bar.grid(row=0, column=1, sticky="ew")

        container.columnconfigure(1, weight=1)
        container.rowconfigure(5, weight=1)
        container.rowconfigure(7, weight=0)

    def _build_playlist_view(self) -> None:
        self.playlist_host_view.columnconfigure(0, weight=1)
        self.playlist_host_view.rowconfigure(0, weight=1)

        self.playlist_tab = PlaylistTab(
            self.playlist_host_view,
            status_callback=lambda text: self.status_text.set(text),
            progress_callback=self._progress_callback,
        )
        self.playlist_tab.frame.grid(row=0, column=0, sticky="nsew")

    def _build_flac_view(self) -> None:
        self.flac_host_view.columnconfigure(0, weight=1)
        self.flac_host_view.rowconfigure(0, weight=1)

        self.flac_converter_tab = FlacConverterTab(
            self.flac_host_view,
            status_callback=lambda text: self.status_text.set(text),
            progress_callback=self._progress_callback,
        )
        self.flac_converter_tab.frame.grid(row=0, column=0, sticky="nsew")

    def _show_view(self, view_name: str) -> None:
        self.active_view.set(view_name)

        self.path_fixer_view.grid_remove()
        self.playlist_host_view.grid_remove()
        self.flac_host_view.grid_remove()

        if view_name == "path_fixer":
            self.path_fixer_view.grid()
            self.root.title("DCSB Config Path Fixer")
        elif view_name == "playlist":
            self.playlist_host_view.grid()
            self.root.title("Playlist Builder")
        elif view_name == "flac_converter":
            self.flac_host_view.grid()
            self.root.title("FLAC -> MP3")
        else:
            return

        self.root.geometry(self.VIEW_SIZES.get(view_name, self.VIEW_SIZES["path_fixer"]))
        self.root.resizable(False, False)
        self._update_tab_button_states()

    def _update_tab_button_states(self) -> None:
        active = self.active_view.get()

        self.path_fixer_tab_button.state(["!disabled"])
        self.playlist_tab_button.state(["!disabled"])
        self.flac_tab_button.state(["!disabled"])

        if active == "path_fixer":
            self.path_fixer_tab_button.state(["disabled"])
        elif active == "playlist":
            self.playlist_tab_button.state(["disabled"])
        elif active == "flac_converter":
            self.flac_tab_button.state(["disabled"])

    def _toggle_config_field(self) -> None:
        state = "normal" if self.use_config.get() else "disabled"
        self.config_entry.config(state=state)
        self.config_browse_btn.config(state=state)

    def _bind_events(self) -> None:
        self.library_path.trace_add("write", self._persist_settings)
        self.config_path.trace_add("write", self._persist_settings)
        self.rename_mode.trace_add("write", self._persist_settings)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _load_settings_into_ui(self) -> None:
        try:
            settings = self.settings.load()
        except SettingsError:
            settings = self.settings.default_settings()

        self.library_path.set(settings["library_path"])
        self.config_path.set(settings["config_path"])

        saved_mode = settings.get("rename_mode", "none")
        self.rename_mode.set(saved_mode if saved_mode in VALID_RENAME_MODES else "none")
        self._toggle_config_field()

    def _persist_settings(self, *_args) -> None:
        try:
            rename_mode = self.rename_mode.get().strip()
            if rename_mode not in VALID_RENAME_MODES:
                rename_mode = "none"

            self.settings.save(self.library_path.get(), self.config_path.get(), rename_mode)
        except SettingsError:
            pass

    def _on_close(self) -> None:
        self._persist_settings()
        self.root.destroy()

    def _browse_library(self) -> None:
        selected = filedialog.askdirectory(title="Select MP3 library folder")
        if selected:
            self.library_path.set(selected)

    def _browse_config(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select DCSB config.xml",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
        )
        if selected:
            self.config_path.set(selected)

    def _clear_output(self) -> None:
        self.log_widget.delete("1.0", tk.END)
        self.missing_widget.delete("1.0", tk.END)
        self.history_details_widget.delete("1.0", tk.END)
        self.status_text.set("Ready.")
        self.progress_value.set(0)

    def _set_log(self, lines: list[str]) -> None:
        self.log_widget.delete("1.0", tk.END)
        if lines:
            self.log_widget.insert("1.0", "\n".join(lines))

    def _set_missing(self, items: list[str]) -> None:
        self.missing_widget.delete("1.0", tk.END)
        if items:
            self.missing_widget.insert("1.0", "\n".join(items))

    def _progress_callback(self, current: int, total: int, message: str) -> None:
        total = max(total, 1)
        percent = max(0, min(100, (current / total) * 100))
        self.progress_value.set(percent)
        self.status_text.set(message)
        self.root.update_idletasks()

    def _refresh_history(self) -> None:
        self.history_listbox.delete(0, tk.END)
        self.history_details_widget.delete("1.0", tk.END)
        self.selected_history_index.set(-1)

        history = self.history_manager.get_history()
        visible_indices: list[int] = []

        for i, entry in enumerate(history):
            entry_type = entry.get("type", "unknown")

            if entry_type == "log":
                continue

            visible_indices.append(i)

            if entry_type == "checkpoint":
                label = entry.get("label", "Untitled Checkpoint")
                display_text = f"[{i}] CHECKPOINT: {label}"
            elif entry_type == "operation":
                count = entry.get("count", len(entry.get("items", [])))
                mode = entry.get("rename_mode", "none")
                lowercase = entry.get("use_lowercase", False)
                extra = " + lowercase" if lowercase else ""
                display_text = f"[{i}] Operation: {count} file(s), mode={mode}{extra}"
            else:
                display_text = f"[{i}] Unknown entry"

            self.history_listbox.insert(tk.END, display_text)

        if not visible_indices:
            self.history_listbox.insert(tk.END, "No history available.")

    def _on_history_select(self, _event) -> None:
        selection = self.history_listbox.curselection()
        if not selection:
            return

        listbox_text = self.history_listbox.get(selection[0])
        if listbox_text == "No history available.":
            return

        try:
            index_text = listbox_text.split("]")[0].lstrip("[")
            index = int(index_text)
        except (ValueError, IndexError):
            return

        history = self.history_manager.get_history()
        if index >= len(history):
            return

        entry = history[index]
        self.selected_history_index.set(index)
        entry_type = entry.get("type", "unknown")

        details = f"Entry #{index}\n"
        details += "━━━━━━━━━━━━━━━━━━━━\n"
        details += f"Timestamp: {entry.get('timestamp', 'N/A')}\n"

        if entry_type == "checkpoint":
            details += "Type: CHECKPOINT\n"
            details += f"Label: {entry.get('label', 'N/A')}\n"
            details += f"Library: {entry.get('library_dir', 'N/A')}\n"
        elif entry_type == "operation":
            items = entry.get("items", [])
            details += "Type: OPERATION\n"
            details += f"Library: {entry.get('library_dir', 'N/A')}\n"
            details += f"Config: {entry.get('config_path', 'N/A')}\n"
            details += f"Mode: {entry.get('rename_mode', 'N/A')}\n"
            details += f"Lowercase: {entry.get('use_lowercase', False)}\n"
            details += f"Renamed files: {len(items)}\n"

            if items:
                details += "\nFirst renamed files:\n"
                for item in items[:10]:
                    old_name = os.path.basename(item.get("old_path", ""))
                    new_name = os.path.basename(item.get("new_path", ""))
                    details += f"- {old_name} -> {new_name}\n"
                if len(items) > 10:
                    details += f"... and {len(items) - 10} more\n"
        else:
            details += f"Type: {entry_type}\n"

        details += "\n━━━━━━━━━━━━━━━━━━━━\n"
        details += f"Revert to this point?\nAll operations after #{index} will be undone."

        self.history_details_widget.delete("1.0", tk.END)
        self.history_details_widget.insert("1.0", details)

    def _revert_to_history(self) -> None:
        history_index = self.selected_history_index.get()

        if history_index < 0:
            messagebox.showwarning("No Selection", "Please select a history entry to revert to.")
            return

        if messagebox.askyesno(
            "Confirm Revert",
            f"Revert all operations after history entry #{history_index}?\nThis action cannot be undone.",
        ):
            self._clear_output()
            self.status_text.set("Reverting...")
            self.progress_value.set(0)

            try:
                result = self.service.revert_to_history_point(history_index, progress_callback=self._progress_callback)
                self._set_log(result.log_lines)
                self._refresh_history()
                self.progress_value.set(100)
                self.status_text.set("Revert complete.")

                messagebox.showinfo(
                    "Revert Complete",
                    f"Reverted: {result.stats.fixed}\nFailed: {result.stats.checked - result.stats.fixed}",
                )
            except Exception as exc:
                self.status_text.set("Revert error.")
                messagebox.showerror("Revert Error", str(exc))

    def _create_checkpoint(self) -> None:
        def on_success(label: str) -> None:
            self._refresh_history()
            self.status_text.set(f"Checkpoint created: {label}")

        open_create_checkpoint_dialog(self.root, self.history_manager, self.library_path.get, on_success)

    def _delete_history_entry(self) -> None:
        history_index = self.selected_history_index.get()

        if history_index < 0:
            messagebox.showwarning("No Selection", "Please select a history entry to delete.")
            return

        history = self.history_manager.get_history()
        if history_index >= len(history):
            messagebox.showerror("Error", "Invalid history entry.")
            return

        entry = history[history_index]
        entry_type = entry.get("type", "unknown")

        if entry_type == "checkpoint":
            display_name = f"Checkpoint: {entry.get('label', 'Untitled')}"
        elif entry_type == "operation":
            display_name = f"Operation: {entry.get('count', len(entry.get('items', [])))} renamed file(s)"
        else:
            display_name = f"Entry type: {entry_type}"

        if messagebox.askyesno(
            "Confirm Delete",
            f"Delete entry #{history_index}?\n{display_name}\n\nThis will remove this entry from history.",
        ):
            try:
                if self.history_manager.delete_entry(history_index):
                    self._refresh_history()
                    self.status_text.set("History entry deleted.")
                    messagebox.showinfo("Deleted", "History entry removed successfully.")
                else:
                    messagebox.showerror("Error", "Failed to delete history entry.")
            except HistoryError as exc:
                self.status_text.set("History error.")
                messagebox.showerror("History Error", str(exc))

    def _rename_checkpoint(self) -> None:
        history_index = self.selected_history_index.get()

        if history_index < 0:
            messagebox.showwarning("No Selection", "Please select a checkpoint to rename.")
            return

        history = self.history_manager.get_history()
        if history_index >= len(history):
            messagebox.showerror("Error", "Invalid history entry.")
            return

        entry = history[history_index]
        if entry.get("type") != "checkpoint":
            messagebox.showwarning("Not a Checkpoint", "You can only rename checkpoint entries.")
            return

        current_label = entry.get("label", "Untitled")

        def on_success(new_label: str) -> None:
            self._refresh_history()
            self.status_text.set(f"Checkpoint renamed to: {new_label}")

        open_rename_checkpoint_dialog(self.root, self.history_manager, history_index, current_label, on_success)

    def _run(self) -> None:
        library_dir = self.library_path.get().strip()
        config_file = self.config_path.get().strip()
        rename_mode = self.rename_mode.get().strip()
        use_lowercase = self.use_lowercase.get()

        if rename_mode not in VALID_RENAME_MODES:
            rename_mode = "none"
            self.rename_mode.set(rename_mode)

        self._clear_output()
        self.status_text.set("Processing...")
        self.progress_value.set(0)

        try:
            self._persist_settings()

            if self.use_config.get():
                result = self.service.repair_config(
                    library_dir=library_dir,
                    config_file=config_file,
                    rename_mode=rename_mode,
                    use_lowercase=use_lowercase,
                    progress_callback=self._progress_callback,
                )
            else:
                result = self.service.normalize_library_only(
                    library_dir=library_dir,
                    rename_mode=rename_mode,
                    use_lowercase=use_lowercase,
                    progress_callback=self._progress_callback,
                )

            self._set_log(result.log_lines)
            self._refresh_history()
            self._set_missing(result.missing_files)
            self.progress_value.set(100)
            self.status_text.set("Done.")

            messagebox.showinfo(
                "Completed",
                f"Checked audio entries: {result.stats.checked}\n"
                f"Fixed paths: {result.stats.fixed}\n"
                f"Missing files: {result.stats.missing}\n"
                f"Ambiguous matches: {result.stats.ambiguous}"
            )

        except ValidationError as exc:
            self.status_text.set("Validation error.")
            messagebox.showerror("Validation Error", str(exc))
        except ConfigReadError as exc:
            self.status_text.set("Read error.")
            messagebox.showerror("Read Error", str(exc))
        except ConfigWriteError as exc:
            self.status_text.set("Write error.")
            messagebox.showerror("Write Error", str(exc))
        except SettingsError as exc:
            self.status_text.set("Settings error.")
            messagebox.showerror("Settings Error", str(exc))
        except Exception as exc:
            self.status_text.set("Unexpected error.")
            messagebox.showerror("Unexpected Error", str(exc))