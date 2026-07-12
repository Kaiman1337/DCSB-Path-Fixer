from __future__ import annotations

import datetime
import json
import os
from pathlib import Path

from core.constants import APP_NAME
from errors import HistoryError


class HistoryManager:
    def __init__(self) -> None:
        self.history_file = self._get_history_file()

    @staticmethod
    def _get_history_file() -> Path:
        appdata = os.getenv("APPDATA")
        base_dir = Path(appdata) if appdata else Path.home() / "AppData" / "Roaming"
        history_dir = base_dir / APP_NAME
        history_dir.mkdir(parents=True, exist_ok=True)
        return history_dir / "rename_history.json"

    def get_history(self) -> list[dict]:
        if not self.history_file.exists():
            return []

        try:
            with self.history_file.open("r", encoding="utf-8") as file:
                data = json.load(file)
                return data if isinstance(data, list) else []
        except (OSError, json.JSONDecodeError):
            return []

    def save_history(self, entries: list[dict]) -> None:
        try:
            with self.history_file.open("w", encoding="utf-8") as file:
                json.dump(entries, file, indent=2, ensure_ascii=False)
        except OSError as exc:
            raise HistoryError(f"Failed to save rename history: {exc}") from exc

    def add_entry(self, old_path: str, new_path: str, library_dir: str, rename_mode: str, use_lowercase: bool) -> None:
        history = self.get_history()
        entry = {
            "timestamp": datetime.datetime.now().isoformat(),
            "type": "rename",
            "old_path": old_path,
            "new_path": new_path,
            "library_dir": library_dir,
            "rename_mode": rename_mode,
            "use_lowercase": use_lowercase,
        }
        history.append(entry)
        self.save_history(history)

    def add_operation_entry(
        self,
        library_dir: str,
        config_path: str,
        rename_mode: str,
        use_lowercase: bool,
        items: list[dict[str, str]],
    ) -> None:
        history = self.get_history()
        entry = {
            "timestamp": datetime.datetime.now().isoformat(),
            "type": "operation",
            "library_dir": library_dir,
            "config_path": config_path,
            "rename_mode": rename_mode,
            "use_lowercase": use_lowercase,
            "items": items,
            "count": len(items),
        }
        history.append(entry)
        self.save_history(history)

    def add_checkpoint(self, label: str, library_dir: str = "") -> None:
        history = self.get_history()
        entry = {
            "timestamp": datetime.datetime.now().isoformat(),
            "type": "checkpoint",
            "label": label,
            "library_dir": library_dir,
        }
        history.append(entry)
        self.save_history(history)

    def get_revert_payload(self, index: int) -> tuple[list[dict[str, str]], list[str]]:
        history = self.get_history()
        if index < 0 or index >= len(history):
            return [], []

        rename_items: list[dict[str, str]] = []
        config_paths: list[str] = []

        for entry in history[index + 1:]:
            entry_type = entry.get("type", "rename")

            if entry_type == "rename":
                old_path = entry.get("old_path")
                new_path = entry.get("new_path")
                if old_path and new_path:
                    rename_items.append({"old_path": old_path, "new_path": new_path})

            elif entry_type == "operation":
                items = entry.get("items", [])
                if isinstance(items, list):
                    for item in items:
                        old_path = item.get("old_path")
                        new_path = item.get("new_path")
                        if old_path and new_path:
                            rename_items.append({"old_path": old_path, "new_path": new_path})

                config_path = str(entry.get("config_path", "")).strip()
                if config_path and config_path not in config_paths:
                    config_paths.append(config_path)

        return rename_items, config_paths

    def clear_history_after_index(self, index: int) -> None:
        history = self.get_history()
        if 0 <= index < len(history):
            history = history[:index + 1]
            self.save_history(history)

    def delete_entry(self, index: int) -> bool:
        history = self.get_history()
        if index < 0 or index >= len(history):
            return False

        entry = history[index]
        history.pop(index)

        try:
            self.save_history(history)
        except HistoryError:
            history.insert(index, entry)
            raise

        entry_type = entry.get("type", "rename")
        if entry_type == "checkpoint":
            self.log_to_history(f"[DELETED] Checkpoint: {entry.get('label', 'Untitled')}")
        elif entry_type == "operation":
            self.log_to_history(f"[DELETED] Operation: {entry.get('count', 0)} file(s)")
        else:
            old_name = Path(entry.get("old_path", "")).name
            new_name = Path(entry.get("new_path", "")).name
            self.log_to_history(f"[DELETED] Rename: {old_name} -> {new_name}")

        return True

    def log_to_history(self, message: str) -> None:
        try:
            history = self.get_history()
            log_entry = {
                "timestamp": datetime.datetime.now().isoformat(),
                "type": "log",
                "message": message,
            }
            history.append(log_entry)
            self.save_history(history)
        except Exception:
            pass

    def rename_checkpoint(self, index: int, new_label: str) -> bool:
        history = self.get_history()
        if index < 0 or index >= len(history):
            return False

        entry = history[index]
        if entry.get("type") != "checkpoint":
            return False

        old_label = entry.get("label", "Untitled")
        entry["label"] = new_label

        try:
            self.save_history(history)
        except HistoryError:
            entry["label"] = old_label
            raise

        self.log_to_history(f"[RENAMED] Checkpoint: {old_label} -> {new_label}")
        return True