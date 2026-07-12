from __future__ import annotations

import os
from pathlib import Path, PureWindowsPath

from core.constants import SUPPORTED_AUDIO_EXTENSIONS
from core.filename_normalizer import generate_name_variants


def is_audio_path(value: str | None) -> bool:
    if not value:
        return False
    return Path(value.strip()).suffix.lower() in SUPPORTED_AUDIO_EXTENSIONS


def normalize_path(path_value: str) -> str:
    return os.path.normpath(path_value.strip())


def normalize_key(path_value: str) -> str:
    return os.path.normcase(os.path.normpath(path_value.strip()))


def generate_path_variants(full_path: str, use_lowercase: bool = False) -> list[str]:
    normalized = normalize_path(full_path)
    path_obj = Path(normalized)
    variants = [normalized]

    for variant_name in generate_name_variants(path_obj.name, use_lowercase):
        variants.append(str(path_obj.with_name(variant_name)))

    unique_variants: list[str] = []
    seen: set[str] = set()

    for variant in variants:
        key = normalize_key(variant)
        if key not in seen:
            seen.add(key)
            unique_variants.append(normalize_path(variant))

    return unique_variants


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


def generate_all_candidates(path_value: str) -> list[str]:
    base_variants = generate_path_variants(path_value)
    all_candidates: list[str] = []

    for variant in base_variants:
        all_candidates.append(variant)
        repaired = generate_repaired_candidates(variant)
        all_candidates.extend(repaired)
        for repaired_variant in repaired:
            all_candidates.extend(generate_path_variants(repaired_variant))

    unique_candidates: list[str] = []
    seen: set[str] = set()

    for candidate in all_candidates:
        key = normalize_key(candidate)
        if key not in seen:
            seen.add(key)
            unique_candidates.append(normalize_path(candidate))

    return unique_candidates


def apply_rename_map(path_value: str, rename_map: dict[str, str]) -> str | None:
    for candidate in generate_all_candidates(path_value):
        key = normalize_key(candidate)
        if key in rename_map:
            return normalize_path(rename_map[key])
    return None


def resolve_path(
    original_path: str,
    audio_index: dict[str, list[str]],
    rename_map: dict[str, str],
) -> tuple[str | None, str, list[str]]:
    normalized_original = normalize_path(original_path)
    all_candidates = generate_all_candidates(normalized_original)

    mapped_path = apply_rename_map(normalized_original, rename_map)
    if mapped_path and os.path.exists(mapped_path):
        return mapped_path, "rename_map", []

    for candidate in all_candidates:
        if os.path.exists(candidate):
            return candidate, "candidate_exists", []

    for candidate in all_candidates:
        key_name = Path(candidate).name.lower()
        matches = audio_index.get(key_name, [])

        if len(matches) == 1:
            return normalize_path(matches[0]), "library_index_variant", []

        if len(matches) > 1:
            return None, "ambiguous_variant", matches

    return None, "not_found", []