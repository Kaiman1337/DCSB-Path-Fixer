from __future__ import annotations

from pathlib import Path

from core.constants import SUPPORTED_AUDIO_EXTENSIONS, ProgressCallback
from core.filename_normalizer import normalize_audio_stem
from core.path_resolver import normalize_key


def collect_audio_files(library_dir: str) -> list[Path]:
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
    library_dir: str,
    progress_callback: ProgressCallback | None = None,
) -> tuple[dict[str, list[str]], int]:
    index: dict[str, list[str]] = {}
    audio_files = collect_audio_files(library_dir)
    total = 0
    max_count = max(len(audio_files), 1)

    for current, file_path in enumerate(audio_files, start=1):
        if progress_callback is not None:
            progress_callback(current, max_count, f"Building audio index... ({current}/{len(audio_files)})")
        index.setdefault(file_path.name.lower(), []).append(str(file_path))
        total += 1

    return index, total


def rename_audio_files(
    library_dir: str,
    replacement: str,
    use_lowercase: bool = False,
    progress_callback: ProgressCallback | None = None,
    log_callback=None,
) -> tuple[int, int, int, dict[str, str], list[dict[str, str]]]:
    renamed = 0
    unchanged = 0
    collisions = 0
    rename_map: dict[str, str] = {}
    renamed_items: list[dict[str, str]] = []

    audio_files = collect_audio_files(library_dir)
    total = max(len(audio_files), 1)

    def log(message: str) -> None:
        if log_callback is not None:
            log_callback(message)

    for index, source in enumerate(audio_files, start=1):
        if progress_callback is not None:
            progress_callback(index, total, f"Renaming files... ({index}/{len(audio_files)})")

        original_stem = source.stem
        suffix = source.suffix
        normalized_stem = normalize_audio_stem(original_stem, replacement, use_lowercase)

        if not normalized_stem or normalized_stem == original_stem:
            unchanged += 1
            continue

        new_name = f"{normalized_stem}{suffix}"
        target = source.with_name(new_name)
        source_key = normalize_key(str(source))

        if str(target) == str(source):
            unchanged += 1
            continue

        if target.exists():
            if str(target).lower() == str(source).lower():
                pass
            else:
                collisions += 1
                rename_map[source_key] = str(target)
                log(f"[SKIP][COLLISION] {source} -> {target}")
                continue

        old_path = str(source)
        source.rename(target)
        new_path = str(target)

        renamed += 1
        rename_map[source_key] = new_path
        renamed_items.append({"old_path": old_path, "new_path": new_path})
        log(f"[RENAMED] {Path(old_path).name} -> {target.name}")

    return renamed, unchanged, collisions, rename_map, renamed_items