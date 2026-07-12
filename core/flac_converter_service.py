from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path

from errors.converter_errors import ConverterDependencyError, ConverterError

APP_TITLE = "FLAC -> MP3 for Winamp"
FFMPEG = "ffmpeg"
FFPROBE = "ffprobe"
BITRATE = "320k"
DURATION_TOLERANCE_SEC = 2


def run_cmd(cmd: list[str]) -> subprocess.CompletedProcess:
    return subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


def tool_exists(name: str) -> bool:
    return shutil.which(name) is not None


def ensure_dependencies() -> None:
    if not tool_exists(FFMPEG):
        raise ConverterDependencyError("ffmpeg not found in PATH.")
    if not tool_exists(FFPROBE):
        raise ConverterDependencyError("ffprobe not found in PATH.")


def ffprobe_duration(path: str) -> float | None:
    cmd = [
        FFPROBE,
        "-v", "error",
        "-show_entries", "format=duration",
        "-of", "default=noprint_wrappers=1:nokey=1",
        path,
    ]
    result = run_cmd(cmd)
    if result.returncode != 0:
        return None

    value = result.stdout.strip().splitlines()
    if not value:
        return None

    try:
        return float(value[0].strip())
    except ValueError:
        return None


def ffprobe_codec(path: str) -> str | None:
    cmd = [
        FFPROBE,
        "-v", "error",
        "-select_streams", "a:0",
        "-show_entries", "stream=codec_name",
        "-of", "default=noprint_wrappers=1:nokey=1",
        path,
    ]
    result = run_cmd(cmd)
    if result.returncode != 0:
        return None

    value = result.stdout.strip().splitlines()
    if not value:
        return None

    return value[0].strip().lower()


def validate_mp3(mp3_path: str, src_duration: float | None) -> tuple[bool, str]:
    if not os.path.isfile(mp3_path):
        return False, "output file missing"

    codec = ffprobe_codec(mp3_path)
    if codec != "mp3":
        return False, f"invalid codec: {codec}"

    out_duration = ffprobe_duration(mp3_path)
    if out_duration is None:
        return False, "cannot read output duration"

    if src_duration is None:
        return False, "cannot read source duration"

    if abs(src_duration - out_duration) > DURATION_TOLERANCE_SEC:
        return False, f"duration mismatch: src={src_duration:.2f}s out={out_duration:.2f}s"

    return True, f"ok: {out_duration:.2f}s"


def convert_one(flac_path: str) -> tuple[bool, str]:
    base, _ = os.path.splitext(flac_path)
    tmp_mp3 = base + ".__tmp__.mp3"
    out_mp3 = base + ".mp3"

    if os.path.exists(tmp_mp3):
        try:
            os.remove(tmp_mp3)
        except OSError:
            return False, f"{os.path.basename(flac_path)} | cannot remove temp file"

    src_duration = ffprobe_duration(flac_path)
    if src_duration is None:
        return False, f"{os.path.basename(flac_path)} | cannot read source duration"

    cmd = [
        FFMPEG,
        "-hide_banner",
        "-loglevel", "error",
        "-nostdin",
        "-y",
        "-i", flac_path,
        "-map_metadata", "0",
        "-map", "0:a:0",
        "-vn",
        "-c:a", "libmp3lame",
        "-b:a", BITRATE,
        "-id3v2_version", "3",
        "-write_id3v1", "1",
        tmp_mp3,
    ]

    result = run_cmd(cmd)
    if result.returncode != 0:
        if os.path.exists(tmp_mp3):
            try:
                os.remove(tmp_mp3)
            except OSError:
                pass
        err = result.stderr.strip() or "ffmpeg encode error"
        return False, f"{os.path.basename(flac_path)} | encode failed | {err}"

    ok, reason = validate_mp3(tmp_mp3, src_duration)
    if not ok:
        if os.path.exists(tmp_mp3):
            try:
                os.remove(tmp_mp3)
            except OSError:
                pass
        return False, f"{os.path.basename(flac_path)} | verify failed | {reason}"

    try:
        if os.path.exists(out_mp3):
            os.remove(out_mp3)
        os.replace(tmp_mp3, out_mp3)
    except OSError as exc:
        if os.path.exists(tmp_mp3):
            try:
                os.remove(tmp_mp3)
            except OSError:
                pass
        return False, f"{os.path.basename(flac_path)} | finalize failed | {exc}"

    try:
        os.remove(flac_path)
    except OSError as exc:
        return False, f"{os.path.basename(flac_path)} | mp3 ok but FLAC delete failed | {exc}"

    return True, f"{os.path.basename(flac_path)} -> {os.path.basename(out_mp3)} | verified"


def list_flac_files_recursive(folder: str) -> list[str]:
    p = Path(folder)
    if not p.is_dir():
        raise ConverterError("Selected path is not a valid folder.")

    flacs: list[str] = []
    for file in p.rglob("*.flac"):
        if file.is_file():
            flacs.append(str(file))
    flacs.sort()
    return flacs


def convert_directory(folder: str, progress_callback=None, log_callback=None) -> tuple[int, int]:
    ensure_dependencies()
    flacs = list_flac_files(folder)

    if not flacs:
        raise ConverterError("No FLAC files found.")

    ok_count = 0
    fail_count = 0

    for idx, flac in enumerate(flacs, start=1):
        if progress_callback is not None:
            progress_callback(idx, len(flacs), f"Processing {idx}/{len(flacs)}: {os.path.basename(flac)}")

        if log_callback is not None:
            log_callback(f"[{idx}/{len(flacs)}] {os.path.basename(flac)}")

        ok, msg = convert_one(flac)

        if log_callback is not None:
            log_callback(("OK    | " if ok else "FAIL  | ") + msg)

        if ok:
            ok_count += 1
        else:
            fail_count += 1

    return ok_count, fail_count