from __future__ import annotations

import os

from core.audio_indexer import build_audio_index, rename_audio_files
from core.config_repair_service import (
    count_xml_audio_entries,
    create_backup,
    parse_config,
    process_element_attributes,
    process_element_text,
    validate_inputs,
    write_config,
)
from core.history_manager import HistoryManager
from core.models import RepairResult, RepairStats
from core.revert_service import revert_to_history_point
from errors import ValidationError


class DCSBPathFixerService:
    def __init__(self) -> None:
        self.log_lines: list[str] = []
        self.history_manager = HistoryManager()

    def log(self, message: str) -> None:
        self.log_lines.append(message)

    @staticmethod
    def _report_progress(callback, current: int, total: int, message: str) -> None:
        if callback is not None:
            callback(current, total, message)

    def repair_config(
        self,
        library_dir: str,
        config_file: str,
        rename_mode: str = "none",
        use_lowercase: bool = False,
        progress_callback=None,
    ) -> RepairResult:
        self.log_lines = []
        validate_inputs(library_dir, config_file)

        self.log(f"Audio library folder: {library_dir}")
        self.log(f"DCSB config file: {config_file}")
        self.log(f"Rename mode: {rename_mode}")
        self.log(f"Convert to lowercase: {'YES' if use_lowercase else 'NO'}")

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
            renamed, unchanged, collisions, rename_map, renamed_items = rename_audio_files(
                library_dir,
                rename_mode,
                use_lowercase,
                progress_callback=lambda c, t, m: self._report_progress(progress_callback, int((c / max(t, 1)) * 40), 100, m),
                log_callback=self.log,
            )
            self.log(f"Rename complete. Renamed: {renamed}, unchanged: {unchanged}, collisions: {collisions}")
        else:
            self._report_progress(progress_callback, 40, 100, "Skipping file renaming...")

        self.history_manager.add_operation_entry(
            library_dir=library_dir,
            config_path=config_file,
            rename_mode=rename_mode,
            use_lowercase=use_lowercase,
            items=renamed_items,
        )
        self.history_manager.log_to_history(f"[OPERATION] Completed operation with {len(renamed_items)} renamed file(s)")

        self.log("Building audio index...")
        audio_index, total_audio_files = build_audio_index(
            library_dir,
            progress_callback=lambda c, t, m: self._report_progress(progress_callback, 40 + int((c / max(t, 1)) * 10), 100, m),
        )
        self.log(f"Indexed audio files: {total_audio_files}")

        self._report_progress(progress_callback, 55, 100, "Creating config backup...")
        backup_path = create_backup(config_file)
        self.log(f"Backup created: {backup_path}")

        self._report_progress(progress_callback, 60, 100, "Reading config...")
        tree = parse_config(config_file)
        root = tree.getroot()

        stats = RepairStats()
        missing_files: list[str] = []
        total_entries = max(count_xml_audio_entries(root), 1)
        processed_entries = 0

        for element in root.iter():
            before_checked = stats.checked
            process_element_text(element, audio_index, rename_map, stats, missing_files, self.log)
            if stats.checked > before_checked:
                processed_entries += stats.checked - before_checked
                percent = 60 + int((processed_entries / total_entries) * 35)
                self._report_progress(progress_callback, min(percent, 95), 100, f"Processing config entries... ({processed_entries}/{total_entries})")

            before_checked = stats.checked
            process_element_attributes(element, audio_index, rename_map, stats, missing_files, self.log)
            if stats.checked > before_checked:
                processed_entries += stats.checked - before_checked
                percent = 60 + int((processed_entries / total_entries) * 35)
                self._report_progress(progress_callback, min(percent, 95), 100, f"Processing config entries... ({processed_entries}/{total_entries})")

        self._report_progress(progress_callback, 98, 100, "Writing config...")
        write_config(tree, config_file)

        self.log("")
        self.log(f"Checked audio entries: {stats.checked}")
        self.log(f"Fixed paths: {stats.fixed}")
        self.log(f"Missing files: {stats.missing}")
        self.log(f"Ambiguous matches: {stats.ambiguous}")

        self._report_progress(progress_callback, 100, 100, "Done.")

        return RepairResult(stats=stats, missing_files=missing_files, log_lines=self.log_lines.copy())

    def normalize_library_only(
        self,
        library_dir: str,
        rename_mode: str = "none",
        use_lowercase: bool = False,
        progress_callback=None,
    ) -> RepairResult:
        self.log_lines = []

        if not library_dir:
            raise ValidationError("MP3 library folder is required.")
        if not os.path.isdir(library_dir):
            raise ValidationError("The MP3 library folder does not exist.")

        self.log(f"Audio library folder: {library_dir}")
        self.log(f"Rename mode: {rename_mode}")
        self.log(f"Convert to lowercase: {'YES' if use_lowercase else 'NO'}")
        self.log("Config file not used - normalizing filenames only.")

        checkpoint_label = f"Starting operation (no config): mode={rename_mode}, lowercase={use_lowercase}"
        self.history_manager.add_checkpoint(checkpoint_label, library_dir)

        renamed, unchanged, collisions, rename_map, renamed_items = rename_audio_files(
            library_dir,
            rename_mode,
            use_lowercase,
            progress_callback=progress_callback,
            log_callback=self.log,
        )
        self.log(f"Rename complete. Renamed: {renamed}, unchanged: {unchanged}, collisions: {collisions}")

        self.history_manager.add_operation_entry(
            library_dir=library_dir,
            config_path="",
            rename_mode=rename_mode,
            use_lowercase=use_lowercase,
            items=renamed_items,
        )
        self.history_manager.log_to_history(f"[OPERATION] Completed operation with {len(renamed_items)} renamed file(s)")

        self._report_progress(progress_callback, 100, 100, "Done.")

        stats = RepairStats(checked=renamed + unchanged, fixed=renamed)
        return RepairResult(stats=stats, missing_files=[], log_lines=self.log_lines.copy())

    def revert_to_history_point(
        self,
        history_index: int,
        progress_callback=None,
    ) -> RepairResult:
        result = revert_to_history_point(self.history_manager, history_index, progress_callback)
        self.log_lines = result.log_lines
        return result