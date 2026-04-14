from __future__ import annotations

import datetime
import json
import os
import re
import shutil
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path, PureWindowsPath
from typing import Callable

from errors import ConfigReadError, ConfigWriteError, HistoryError, SettingsError, ValidationError

APP_NAME = "DCSBPathFixer"
SUPPORTED_AUDIO_EXTENSIONS = {
    ".mp3",
    ".wav",
    ".flac",
    ".ogg",
    ".oga",
    ".m4a",
    ".aac",
    ".wma",
    ".opus",
    ".aiff",
    ".aif",
    ".ape",
    ".alac",
    ".mp2",
    ".mpga",
    ".m4b",
}

ProgressCallback = Callable[[int, int, str], None]


@dataclass(slots=True)
class RepairStats:
    checked: int = 0
    fixed: int = 0
    missing: int = 0
    ambiguous: int = 0


@dataclass(slots=True)
class RepairResult:
    stats: RepairStats
    missing_files: list[str]
    log_lines: list[str]


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

class DCSBPathFixerService:
    def __init__(self) -> None:
        self.log_lines: list[str] = []
        self.history_manager = HistoryManager()

    def log(self, message: str) -> None:
        self.log_lines.append(message)

    @staticmethod
    def _report_progress(callback: ProgressCallback | None, current: int, total: int, message: str) -> None:
        if callback is not None:
            callback(current, total, message)

    def validate_inputs(self, library_dir: str, config_file: str) -> None:
        if not library_dir:
            raise ValidationError("MP3 library folder is required.")
        if not config_file:
            raise ValidationError("DCSB config.xml file is required.")
        if not os.path.isdir(library_dir):
            raise ValidationError("The MP3 library folder does not exist.")
        if not os.path.isfile(config_file):
            raise ValidationError("The config.xml file does not exist.")

    @staticmethod
    def is_audio_path(value: str | None) -> bool:
        if not value:
            return False
        return Path(value.strip()).suffix.lower() in SUPPORTED_AUDIO_EXTENSIONS

    @staticmethod
    def normalize_path(path_value: str) -> str:
        return os.path.normpath(path_value.strip())

    @staticmethod
    def normalize_key(path_value: str) -> str:
        return os.path.normcase(os.path.normpath(path_value.strip()))

    @staticmethod
    def create_backup(config_file: str) -> str:
        source = Path(config_file)
        backup = source.with_name(f"{source.stem}.backup{source.suffix}")
        shutil.copy2(source, backup)
        return str(backup)

    @staticmethod
    def parse_config(config_file: str) -> ET.ElementTree:
        try:
            return ET.parse(config_file)
        except (ET.ParseError, OSError) as exc:
            raise ConfigReadError(f"Failed to read XML config: {exc}") from exc

    @staticmethod
    def write_config(tree: ET.ElementTree, config_file: str) -> None:
        try:
            tree.write(config_file, encoding="utf-8", xml_declaration=True)
        except OSError as exc:
            raise ConfigWriteError(f"Failed to write XML config: {exc}") from exc

    @staticmethod
    def normalize_audio_stem(stem: str, rename_mode: str, use_lowercase: bool = False) -> str:
        original_text = stem.strip()

        if rename_mode == "none":
            text = original_text
        else:
            text = original_text
            text = text.replace("–", "-").replace("—", "-")
            text = re.sub(r"_", " ", text)
            text = re.sub(r"-", " ", text)
            text = re.sub(r"\s+", " ", text)

            text = re.sub(r"([!?.,;:])\s*(?:\1\s*)+", r"\1", text)
            text = re.sub(r"\(\s*(\d+)\s*[-\s]*kbps\s*\)", r"(\1kbps)", text, flags=re.IGNORECASE)
            text = re.sub(r"\(\s*(\d+)\s*[-\s]*bit\s*\)", r"(\1bit)", text, flags=re.IGNORECASE)
            text = re.sub(r"\s*\(\d+\)\s*$", "", text)
            text = re.sub(r"\b(featuring|feat|ft)\s*\.?\s*", "ft.", text, flags=re.IGNORECASE)
            text = re.sub(r"\bversus\b", "vs.", text, flags=re.IGNORECASE)
            text = re.sub(r"\bvs\s*\.?\s*", "vs.", text, flags=re.IGNORECASE)
            text = re.sub(r"\s*ft\.\s*", " ft. ", text, flags=re.IGNORECASE)
            text = re.sub(r"\s*vs\.\s*", " vs. ", text, flags=re.IGNORECASE)
            text = re.sub(r"\s*([!?,;:])\s*", r"\1 ", text)
            text = re.sub(r"\b0+(\d+)\b", r"\1", text)
            text = text.strip(" .-_")
            text = re.sub(r"\s{2,}", " ", text)

            if rename_mode == "space":
                pass  # text is already normalized to spaces
            elif rename_mode == "-":
                text = re.sub(r"\s*ft\.\s*", "-ft.", text, flags=re.IGNORECASE)
                text = re.sub(r"\s*vs\.\s*", "-vs.", text, flags=re.IGNORECASE)
                text = re.sub(r"\s+", "-", text)
                text = re.sub(r"-{2,}", "-", text)
                text = text.strip("-")
            elif rename_mode == "_":
                text = re.sub(r"\s*ft\.\s*", "_ft.", text, flags=re.IGNORECASE)
                text = re.sub(r"\s*vs\.\s*", "_vs.", text, flags=re.IGNORECASE)
                text = re.sub(r"\s+", "_", text)
                text = re.sub(r"_{2,}", "_", text)
                text = text.strip("_")

        # Apply lowercase conversion independently of rename mode
        if use_lowercase:
            text = text.lower()

        return text

    def rename_audio_files(
        self,
        library_dir: str,
        replacement: str,
        use_lowercase: bool = False,
        progress_callback: ProgressCallback | None = None,
    ) -> tuple[int, int, int, dict[str, str], list[dict[str, str]]]:
        renamed = 0
        unchanged = 0
        collisions = 0
        rename_map: dict[str, str] = {}
        renamed_items: list[dict[str, str]] = []

        audio_files = self._collect_audio_files(library_dir)
        total = max(len(audio_files), 1)

        for index, source in enumerate(audio_files, start=1):
            self._report_progress(progress_callback, index, total, f"Renaming files... ({index}/{len(audio_files)})")

            original_stem = source.stem
            suffix = source.suffix
            normalized_stem = self.normalize_audio_stem(original_stem, replacement, use_lowercase)

            if not normalized_stem or normalized_stem == original_stem:
                unchanged += 1
                continue

            new_name = f"{normalized_stem}{suffix}"
            target = source.with_name(new_name)
            source_key = self.normalize_key(str(source))

            # Check if target path is exactly the same as source (including case)
            if str(target) == str(source):
                unchanged += 1
                continue

            if target.exists():
                # On case-insensitive filesystems, check if it's the same file (different case)
                if str(target).lower() == str(source).lower():
                    # Same file, different case - allow rename for case change
                    pass
                else:
                    # Different file, collision
                    collisions += 1
                    rename_map[source_key] = str(target)
                    self.log(f"[SKIP][COLLISION] {source} -> {target}")
                    continue

            old_path = str(source)
            source.rename(target)
            new_path = str(target)

            renamed += 1
            rename_map[source_key] = new_path
            renamed_items.append({"old_path": old_path, "new_path": new_path})
            self.log(f"[RENAMED] {Path(old_path).name} -> {target.name}")

        return renamed, unchanged, collisions, rename_map, renamed_items

    def _collect_audio_files(self, library_dir: str) -> list[Path]:
        library = Path(library_dir)
        if not library.is_dir():
            return []

        files = []
        for file in library.rglob("*"):
            if file.is_file() and file.suffix.lower() in SUPPORTED_AUDIO_EXTENSIONS:
                files.append(file)

        files.sort(key=lambda p: str(p).lower())
        return files

    def build_audio_index(
        self,
        library_dir: str,
        progress_callback: ProgressCallback | None = None,
    ) -> tuple[dict[str, list[str]], int]:
        index: dict[str, list[str]] = {}
        audio_files = self._collect_audio_files(library_dir)
        total = 0
        max_count = max(len(audio_files), 1)

        for current, file_path in enumerate(audio_files, start=1):
            self._report_progress(progress_callback, current, max_count, f"Building audio index... ({current}/{len(audio_files)})")
            index.setdefault(file_path.name.lower(), []).append(str(file_path))
            total += 1

        return index, total

    @staticmethod
    def generate_name_variants(filename: str, use_lowercase: bool = False) -> list[str]:
        path = Path(filename)
        stem = path.stem
        suffix = path.suffix

        normalized_space = DCSBPathFixerService.normalize_audio_stem(stem, "space", use_lowercase)
        normalized_dash = DCSBPathFixerService.normalize_audio_stem(stem, "-", use_lowercase)
        normalized_underscore = DCSBPathFixerService.normalize_audio_stem(stem, "_", use_lowercase)
        normalized_none = DCSBPathFixerService.normalize_audio_stem(stem, "none", use_lowercase)

        variants = [
            f"{normalized_space}{suffix}",
            f"{normalized_dash}{suffix}",
            f"{normalized_underscore}{suffix}",
            f"{normalized_none}{suffix}",
        ]

        unique_variants: list[str] = []
        seen: set[str] = set()

        for variant in variants:
            key = variant.lower()
            if key not in seen:
                seen.add(key)
                unique_variants.append(variant)

        return unique_variants

    def generate_path_variants(self, full_path: str, use_lowercase: bool = False) -> list[str]:
        normalized = self.normalize_path(full_path)
        path_obj = Path(normalized)
        variants = [normalized]

        for variant_name in self.generate_name_variants(path_obj.name, use_lowercase):
            variants.append(str(path_obj.with_name(variant_name)))

        unique_variants: list[str] = []
        seen: set[str] = set()

        for variant in variants:
            key = self.normalize_key(variant)
            if key not in seen:
                seen.add(key)
                unique_variants.append(self.normalize_path(variant))

        return unique_variants

    @staticmethod
    def generate_repaired_candidates(path_value: str) -> list[str]:
        raw = path_value.strip()
        if not raw:
            return []

        path_obj = PureWindowsPath(raw)
        parent = path_obj.parent
        filename = path_obj.name
        folder_name = parent.name

        if not folder_name or not filename:
            return []

        filename_lower = filename.lower()
        folder_lower = folder_name.lower()

        if not filename_lower.startswith(folder_lower):
            return []

        trimmed_name = filename[len(folder_name):].lstrip(" _-")
        if not trimmed_name:
            return []

        candidates = [
            os.path.normpath(str(parent / trimmed_name)),
            os.path.normpath(str(parent / folder_name / trimmed_name)),
            os.path.normpath(str(parent / folder_name / filename)),
        ]

        unique_candidates: list[str] = []
        seen: set[str] = set()

        for candidate in candidates:
            key = os.path.normcase(os.path.normpath(candidate))
            if key not in seen:
                seen.add(key)
                unique_candidates.append(candidate)

        return unique_candidates

    def generate_all_candidates(self, path_value: str) -> list[str]:
        base_variants = self.generate_path_variants(path_value)
        all_candidates: list[str] = []

        for variant in base_variants:
            all_candidates.append(variant)
            repaired = self.generate_repaired_candidates(variant)
            all_candidates.extend(repaired)
            for repaired_variant in repaired:
                all_candidates.extend(self.generate_path_variants(repaired_variant))

        unique_candidates: list[str] = []
        seen: set[str] = set()

        for candidate in all_candidates:
            key = self.normalize_key(candidate)
            if key not in seen:
                seen.add(key)
                unique_candidates.append(self.normalize_path(candidate))

        return unique_candidates

    def apply_rename_map(self, path_value: str, rename_map: dict[str, str]) -> str | None:
        for candidate in self.generate_all_candidates(path_value):
            key = self.normalize_key(candidate)
            if key in rename_map:
                return self.normalize_path(rename_map[key])
        return None

    def resolve_path(
        self,
        original_path: str,
        audio_index: dict[str, list[str]],
        rename_map: dict[str, str],
    ) -> tuple[str | None, str, list[str]]:
        normalized_original = self.normalize_path(original_path)
        all_candidates = self.generate_all_candidates(normalized_original)

        mapped_path = self.apply_rename_map(normalized_original, rename_map)
        if mapped_path and os.path.exists(mapped_path):
            return mapped_path, "rename_map", []

        for candidate in all_candidates:
            if os.path.exists(candidate):
                return candidate, "candidate_exists", []

        for candidate in all_candidates:
            key_name = Path(candidate).name.lower()
            matches = audio_index.get(key_name, [])

            if len(matches) == 1:
                return self.normalize_path(matches[0]), "library_index_variant", []

            if len(matches) > 1:
                return None, "ambiguous_variant", matches

        return None, "not_found", []

    def process_element_text(
        self,
        element: ET.Element,
        audio_index: dict[str, list[str]],
        rename_map: dict[str, str],
        stats: RepairStats,
        missing_files: list[str],
    ) -> None:
        if not element.text or not self.is_audio_path(element.text):
            return

        original_path = element.text.strip()
        stats.checked += 1

        resolved_path, reason, matches = self.resolve_path(original_path, audio_index, rename_map)

        if resolved_path:
            if self.normalize_key(resolved_path) != self.normalize_key(original_path):
                element.text = resolved_path
                stats.fixed += 1
                self.log(f"[FIXED][{reason.upper()}] {original_path} -> {resolved_path}")
            else:
                self.log(f"[OK] {original_path}")
            return

        if reason.startswith("ambiguous"):
            stats.ambiguous += 1
            self.log(f"[AMBIGUOUS] {original_path}")
            for match in matches:
                self.log(f" -> {match}")
            return

        stats.missing += 1
        missing_files.append(original_path)
        self.log(f"[NOT FOUND] {original_path}")

    def process_element_attributes(
        self,
        element: ET.Element,
        audio_index: dict[str, list[str]],
        rename_map: dict[str, str],
        stats: RepairStats,
        missing_files: list[str],
    ) -> None:
        for attr_name, attr_value in element.attrib.items():
            if not self.is_audio_path(attr_value):
                continue

            original_path = attr_value.strip()
            stats.checked += 1

            resolved_path, reason, matches = self.resolve_path(original_path, audio_index, rename_map)

            if resolved_path:
                if self.normalize_key(resolved_path) != self.normalize_key(original_path):
                    element.attrib[attr_name] = resolved_path
                    stats.fixed += 1
                    self.log(f"[FIXED][ATTR][{reason.upper()}] {original_path} -> {resolved_path}")
                else:
                    self.log(f"[OK][ATTR] {original_path}")
                continue

            if reason.startswith("ambiguous"):
                stats.ambiguous += 1
                self.log(f"[AMBIGUOUS][ATTR] {original_path}")
                for match in matches:
                    self.log(f" -> {match}")
                continue

            stats.missing += 1
            missing_files.append(original_path)
            self.log(f"[NOT FOUND][ATTR] {original_path}")

    @staticmethod
    def _count_xml_audio_entries(root: ET.Element) -> int:
        total = 0
        for element in root.iter():
            if element.text and DCSBPathFixerService.is_audio_path(element.text):
                total += 1
            for attr_value in element.attrib.values():
                if DCSBPathFixerService.is_audio_path(attr_value):
                    total += 1
        return total

    def repair_config(
        self,
        library_dir: str,
        config_file: str,
        rename_mode: str = "none",
        use_lowercase: bool = False,
        progress_callback: ProgressCallback | None = None,
    ) -> RepairResult:
        self.log_lines = []
        self.validate_inputs(library_dir, config_file)

        self.log(f"Audio library folder: {library_dir}")
        self.log(f"DCSB config file: {config_file}")
        self.log(f"Rename mode: {rename_mode}")
        self.log(f"Convert to lowercase: {'YES' if use_lowercase else 'NO'}")

        # Add an operation checkpoint before performing file updates
        checkpoint_label = f"Starting operation: mode={rename_mode}, lowercase={use_lowercase}"
        self.history_manager.add_checkpoint(checkpoint_label, library_dir)

        rename_map: dict[str, str] = {}
        renamed_items: list[dict[str, str]] = []

        if rename_mode in {"-", "_", "space"} and use_lowercase:
            self.log("Filename normalization and lowercase conversion enabled.")
        elif rename_mode in {"-", "_", "space"}:
            self.log("Filename normalization is enabled.")
        elif use_lowercase:
            self.log("Lowercase conversion is enabled.")
        else:
            self.log("No filename changes selected.")

        if rename_mode != "none" or use_lowercase:
            self.log("Renaming audio files...")
            renamed, unchanged, collisions, rename_map, renamed_items = self.rename_audio_files(
                library_dir,
                rename_mode,
                use_lowercase,
                progress_callback=lambda c, t, m: self._report_progress(progress_callback, int((c / max(t, 1)) * 40), 100, m),
            )
            self.log(f"Rename complete. Renamed: {renamed}, unchanged: {unchanged}, collisions: {collisions}")
        else:
            self._report_progress(progress_callback, 40, 100, "Skipping file renaming...")

        # Save a single operation entry for the whole run (not each file)
        self.history_manager.add_operation_entry(
            library_dir=library_dir,
            config_path=config_file,
            rename_mode=rename_mode,
            use_lowercase=use_lowercase,
            items=renamed_items,
        )
        self.history_manager.log_to_history(f"[OPERATION] Completed operation with {len(renamed_items)} renamed file(s)")

        self.log("Building audio index...")
        audio_index, total_audio_files = self.build_audio_index(
            library_dir,
            progress_callback=lambda c, t, m: self._report_progress(progress_callback, 40 + int((c / max(t, 1)) * 10), 100, m),
        )
        self.log(f"Indexed audio files: {total_audio_files}")

        self._report_progress(progress_callback, 55, 100, "Creating config backup...")
        backup_path = self.create_backup(config_file)
        self.log(f"Backup created: {backup_path}")

        self._report_progress(progress_callback, 60, 100, "Reading config...")
        tree = self.parse_config(config_file)
        root = tree.getroot()

        stats = RepairStats()
        missing_files: list[str] = []
        total_entries = max(self._count_xml_audio_entries(root), 1)
        processed_entries = 0

        for element in root.iter():
            before_checked = stats.checked
            self.process_element_text(element, audio_index, rename_map, stats, missing_files)
            if stats.checked > before_checked:
                processed_entries += stats.checked - before_checked
                percent = 60 + int((processed_entries / total_entries) * 35)
                self._report_progress(progress_callback, min(percent, 95), 100, f"Processing config entries... ({processed_entries}/{total_entries})")

            before_checked = stats.checked
            self.process_element_attributes(element, audio_index, rename_map, stats, missing_files)
            if stats.checked > before_checked:
                processed_entries += stats.checked - before_checked
                percent = 60 + int((processed_entries / total_entries) * 35)
                self._report_progress(progress_callback, min(percent, 95), 100, f"Processing config entries... ({processed_entries}/{total_entries})")

        self._report_progress(progress_callback, 98, 100, "Writing config...")
        self.write_config(tree, config_file)

        self.log("")
        self.log(f"Checked audio entries: {stats.checked}")
        self.log(f"Fixed paths: {stats.fixed}")
        self.log(f"Missing files: {stats.missing}")
        self.log(f"Ambiguous matches: {stats.ambiguous}")

        self._report_progress(progress_callback, 100, 100, "Done.")

        return RepairResult(
            stats=stats,
            missing_files=missing_files,
            log_lines=self.log_lines.copy(),
        )

    def _revert_paths_in_config(self, config_file: str, rename_items: list[dict[str, str]]) -> int:
        if not config_file or not os.path.isfile(config_file):
            return 0

        tree = self.parse_config(config_file)
        root = tree.getroot()

        replacements = {
            self.normalize_key(item["new_path"]): self.normalize_path(item["old_path"])
            for item in rename_items
        }

        changed = 0

        for element in root.iter():
            if element.text and self.is_audio_path(element.text):
                key = self.normalize_key(element.text)
                if key in replacements:
                    element.text = replacements[key]
                    changed += 1

            for attr_name, attr_value in list(element.attrib.items()):
                if self.is_audio_path(attr_value):
                    key = self.normalize_key(attr_value)
                    if key in replacements:
                        element.attrib[attr_name] = replacements[key]
                        changed += 1

        if changed > 0:
            self.write_config(tree, config_file)

        return changed

    def revert_to_history_point(
        self,
        history_index: int,
        progress_callback: ProgressCallback | None = None,
    ) -> RepairResult:
        self.log_lines = []

        self._report_progress(progress_callback, 0, 100, "Revert started...")

        rename_items, config_paths = self.history_manager.get_revert_payload(history_index)

        if not rename_items:
            self.log("No operations to revert.")
            self.history_manager.log_to_history(f"[REVERT] No operations to revert after index {history_index}")
            return RepairResult(
                stats=RepairStats(),
                missing_files=[],
                log_lines=self.log_lines.copy(),
            )

        self.log(f"Reverting {len(rename_items)} file rename(s)...")

        reverted = 0
        failed = 0

        total_files = len(rename_items)

        for idx, item in enumerate(reversed(rename_items), start=1):
            old_path = item["old_path"]
            new_path = item["new_path"]

            try:
                new_file = Path(new_path)
                old_file = Path(old_path)

                if new_file.exists():
                    old_file.parent.mkdir(parents=True, exist_ok=True)
                    new_file.rename(old_file)
                    reverted += 1
                    self.log(f"[REVERTED] {new_file.name} -> {old_file.name}")
                    self.history_manager.log_to_history(f"[REVERTED] {new_file.name} -> {old_file.name}")
                else:
                    failed += 1
                    self.log(f"[REVERT FAILED] File not found: {new_path}")
                    self.history_manager.log_to_history(f"[REVERT FAILED] File not found: {Path(new_path).name}")
            except Exception as exc:
                failed += 1
                self.log(f"[REVERT ERROR] {new_path}: {exc}")
                self.history_manager.log_to_history(f"[REVERT ERROR] {Path(new_path).name}: {exc}")

            self._report_progress(progress_callback, int((idx / max(total_files, 1)) * 50), 100, f"Reverting files... ({idx}/{total_files})")

        config_updates = 0
        for idx, config_path in enumerate(config_paths, start=1):
            try:
                changed = self._revert_paths_in_config(config_path, rename_items)
                config_updates += changed
                self.log(f"[CONFIG UPDATED] {config_path} ({changed} path change(s))")
            except Exception as exc:
                self.log(f"[CONFIG UPDATE ERROR] {config_path}: {exc}")

            self._report_progress(
                progress_callback,
                50 + int((idx / max(len(config_paths), 1)) * 40),
                100,
                f"Reverting config paths... ({idx}/{len(config_paths)})",
            )

        self.history_manager.clear_history_after_index(history_index)

        self.log("")
        self.log(f"Revert complete. Reverted: {reverted}, Failed: {failed}")
        self.log(f"Config paths updated: {config_updates}")

        self._report_progress(progress_callback, 100, 100, "Revert complete.")

        return RepairResult(
            stats=RepairStats(checked=len(rename_items), fixed=reverted),
            missing_files=[],
            log_lines=self.log_lines.copy(),
        )