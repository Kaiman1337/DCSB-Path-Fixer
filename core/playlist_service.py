from __future__ import annotations

import os
from pathlib import Path

from errors.playlist_errors import PlaylistError


SUPPORTED_PLAYLIST_AUDIO_EXTENSIONS = {
    ".mp3", ".wav", ".flac", ".ogg", ".oga", ".m4a", ".aac",
    ".wma", ".opus", ".aiff", ".aif", ".ape", ".alac",
    ".mp2", ".mpga", ".m4b",
}


def is_audio_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in SUPPORTED_PLAYLIST_AUDIO_EXTENSIONS


def collect_audio_files(folder: str, recursive: bool = True) -> list[Path]:
    root = Path(folder)

    if not root.is_dir():
        raise PlaylistError("Selected path is not a valid folder.")

    if recursive:
        files = [p for p in root.rglob("*") if is_audio_file(p)]
    else:
        files = [p for p in root.iterdir() if is_audio_file(p)]

    files.sort(key=lambda p: str(p).lower())
    return files


def build_m3u_content(audio_files: list[Path], playlist_name: str | None = None) -> str:
    lines = ["#EXTM3U"]

    if playlist_name:
        lines.append(f"#PLAYLIST:{playlist_name}")

    for file_path in audio_files:
        lines.append(str(file_path))

    return "\n".join(lines) + "\n"


def create_playlist(
    source_folder: str,
    output_path: str | None = None,
    recursive: bool = True,
    progress_callback=None,
    log_callback=None,
) -> tuple[str, int]:
    audio_files = collect_audio_files(source_folder, recursive=recursive)

    if not audio_files:
        raise PlaylistError("No audio files found.")

    source = Path(source_folder)

    if output_path:
        playlist_path = Path(output_path)
    else:
        playlist_path = source / f"{source.name}.m3u"

    playlist_name = playlist_path.stem

    total = len(audio_files)
    for idx, file_path in enumerate(audio_files, start=1):
        if progress_callback is not None:
            progress_callback(idx, total, f"Adding files to playlist... ({idx}/{total})")

        if log_callback is not None:
            log_callback(f"[{idx}/{total}] {file_path.name}")

    content = build_m3u_content(audio_files, playlist_name=playlist_name)

    try:
        playlist_path.parent.mkdir(parents=True, exist_ok=True)
        playlist_path.write_text(content, encoding="utf-8", newline="\n")
    except OSError as exc:
        raise PlaylistError(f"Failed to write playlist file: {exc}") from exc

    return str(playlist_path), len(audio_files)