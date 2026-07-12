from __future__ import annotations

import os
from pathlib import Path

from core.config_repair_service import parse_config, write_config
from core.history_manager import HistoryManager
from core.models import RepairResult, RepairStats
from core.path_resolver import is_audio_path, normalize_key, normalize_path


def revert_paths_in_config(config_file: str, rename_items: list[dict[str, str]]) -> int:
    if not config_file or not os.path.isfile(config_file):
        return 0

    tree = parse_config(config_file)
    root = tree.getroot()

    replacements = {
        normalize_key(item["new_path"]): normalize_path(item["old_path"])
        for item in rename_items
    }

    changed = 0

    for element in root.iter():
        if element.text and is_audio_path(element.text):
            key = normalize_key(element.text)
            if key in replacements:
                element.text = replacements[key]
                changed += 1

        for attr_name, attr_value in list(element.attrib.items()):
            if is_audio_path(attr_value):
                key = normalize_key(attr_value)
                if key in replacements:
                    element.attrib[attr_name] = replacements[key]
                    changed += 1

    if changed > 0:
        write_config(tree, config_file)

    return changed


def revert_to_history_point(
    history_manager: HistoryManager,
    history_index: int,
    progress_callback=None,
) -> RepairResult:
    log_lines: list[str] = []

    def log(message: str) -> None:
        log_lines.append(message)

    def report(current: int, total: int, message: str) -> None:
        if progress_callback is not None:
            progress_callback(current, total, message)

    report(0, 100, "Revert started...")

    rename_items, config_paths = history_manager.get_revert_payload(history_index)

    if not rename_items:
        log("No operations to revert.")
        history_manager.log_to_history(f"[REVERT] No operations to revert after index {history_index}")
        return RepairResult(stats=RepairStats(), missing_files=[], log_lines=log_lines.copy())

    log(f"Reverting {len(rename_items)} file rename(s)...")

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
                log(f"[REVERTED] {new_file.name} -> {old_file.name}")
                history_manager.log_to_history(f"[REVERTED] {new_file.name} -> {old_file.name}")
            else:
                failed += 1
                log(f"[REVERT FAILED] File not found: {new_path}")
                history_manager.log_to_history(f"[REVERT FAILED] File not found: {Path(new_path).name}")
        except Exception as exc:
            failed += 1
            log(f"[REVERT ERROR] {new_path}: {exc}")
            history_manager.log_to_history(f"[REVERT ERROR] {Path(new_path).name}: {exc}")

        report(int((idx / max(total_files, 1)) * 50), 100, f"Reverting files... ({idx}/{total_files})")

    config_updates = 0
    for idx, config_path in enumerate(config_paths, start=1):
        try:
            changed = revert_paths_in_config(config_path, rename_items)
            config_updates += changed
            log(f"[CONFIG UPDATED] {config_path} ({changed} path change(s))")
        except Exception as exc:
            log(f"[CONFIG UPDATE ERROR] {config_path}: {exc}")

        report(
            50 + int((idx / max(len(config_paths), 1)) * 40),
            100,
            f"Reverting config paths... ({idx}/{len(config_paths)})",
        )

    history_manager.clear_history_after_index(history_index)

    log("")
    log(f"Revert complete. Reverted: {reverted}, Failed: {failed}")
    log(f"Config paths updated: {config_updates}")

    report(100, 100, "Revert complete.")

    return RepairResult(
        stats=RepairStats(checked=len(rename_items), fixed=reverted),
        missing_files=[],
        log_lines=log_lines.copy(),
    )