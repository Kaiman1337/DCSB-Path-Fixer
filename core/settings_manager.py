from __future__ import annotations

import json
import os
from pathlib import Path

from core.constants import APP_NAME
from errors import SettingsError


class SettingsManager:
    def __init__(self) -> None:
        self.settings_file = self._get_settings_file()

    @staticmethod
    def _get_settings_file() -> Path:
        appdata = os.getenv("APPDATA")
        base_dir = Path(appdata) if appdata else Path.home() / "AppData" / "Roaming"
        settings_dir = base_dir / APP_NAME
        settings_dir.mkdir(parents=True, exist_ok=True)
        return settings_dir / "settings.json"

    @staticmethod
    def default_settings() -> dict[str, str]:
        return {
            "library_path": "",
            "config_path": "",
            "rename_mode": "none",
        }

    def load(self) -> dict[str, str]:
        if not self.settings_file.exists():
            return self.default_settings()

        try:
            with self.settings_file.open("r", encoding="utf-8") as file:
                data = json.load(file)
        except (OSError, json.JSONDecodeError) as exc:
            raise SettingsError(f"Failed to load settings: {exc}") from exc

        settings = self.default_settings()
        settings.update(
            {
                "library_path": str(data.get("library_path", "")),
                "config_path": str(data.get("config_path", "")),
                "rename_mode": str(data.get("rename_mode", "none")),
            }
        )
        return settings

    def save(self, library_path: str, config_path: str, rename_mode: str) -> None:
        payload = {
            "library_path": library_path.strip(),
            "config_path": config_path.strip(),
            "rename_mode": rename_mode.strip() or "none",
        }

        try:
            with self.settings_file.open("w", encoding="utf-8") as file:
                json.dump(payload, file, indent=2, ensure_ascii=False)
        except OSError as exc:
            raise SettingsError(f"Failed to save settings: {exc}") from exc