from __future__ import annotations

import os
import shutil
import xml.etree.ElementTree as ET
from pathlib import Path

from core.models import RepairStats
from core.path_resolver import is_audio_path, normalize_key, resolve_path
from errors import ConfigReadError, ConfigWriteError, ValidationError


def validate_inputs(library_dir: str, config_file: str) -> None:
    if not library_dir:
        raise ValidationError("MP3 library folder is required.")
    if not config_file:
        raise ValidationError("DCSB config.xml file is required.")
    if not os.path.isdir(library_dir):
        raise ValidationError("The MP3 library folder does not exist.")
    if not os.path.isfile(config_file):
        raise ValidationError("The config.xml file does not exist.")


def create_backup(config_file: str) -> str:
    source = Path(config_file)
    backup = source.with_name(f"{source.stem}.backup{source.suffix}")
    shutil.copy2(source, backup)
    return str(backup)


def parse_config(config_file: str) -> ET.ElementTree:
    try:
        return ET.parse(config_file)
    except (ET.ParseError, OSError) as exc:
        raise ConfigReadError(f"Failed to read XML config: {exc}") from exc


def write_config(tree: ET.ElementTree, config_file: str) -> None:
    try:
        tree.write(config_file, encoding="utf-8", xml_declaration=True)
    except OSError as exc:
        raise ConfigWriteError(f"Failed to write XML config: {exc}") from exc


def count_xml_audio_entries(root: ET.Element) -> int:
    total = 0
    for element in root.iter():
        if element.text and is_audio_path(element.text):
            total += 1
        for attr_value in element.attrib.values():
            if is_audio_path(attr_value):
                total += 1
    return total


def process_element_text(
    element: ET.Element,
    audio_index: dict[str, list[str]],
    rename_map: dict[str, str],
    stats: RepairStats,
    missing_files: list[str],
    log_callback,
) -> None:
    if not element.text or not is_audio_path(element.text):
        return

    original_path = element.text.strip()
    stats.checked += 1

    resolved_path, reason, matches = resolve_path(original_path, audio_index, rename_map)

    if resolved_path:
        if normalize_key(resolved_path) != normalize_key(original_path):
            element.text = resolved_path
            stats.fixed += 1
            log_callback(f"[FIXED][{reason.upper()}] {original_path} -> {resolved_path}")
        else:
            log_callback(f"[OK] {original_path}")
        return

    if reason.startswith("ambiguous"):
        stats.ambiguous += 1
        log_callback(f"[AMBIGUOUS] {original_path}")
        for match in matches:
            log_callback(f" -> {match}")
        return

    stats.missing += 1
    missing_files.append(original_path)
    log_callback(f"[NOT FOUND] {original_path}")


def process_element_attributes(
    element: ET.Element,
    audio_index: dict[str, list[str]],
    rename_map: dict[str, str],
    stats: RepairStats,
    missing_files: list[str],
    log_callback,
) -> None:
    for attr_name, attr_value in element.attrib.items():
        if not is_audio_path(attr_value):
            continue

        original_path = attr_value.strip()
        stats.checked += 1

        resolved_path, reason, matches = resolve_path(original_path, audio_index, rename_map)

        if resolved_path:
            if normalize_key(resolved_path) != normalize_key(original_path):
                element.attrib[attr_name] = resolved_path
                stats.fixed += 1
                log_callback(f"[FIXED][ATTR][{reason.upper()}] {original_path} -> {resolved_path}")
            else:
                log_callback(f"[OK][ATTR] {original_path}")
            continue

        if reason.startswith("ambiguous"):
            stats.ambiguous += 1
            log_callback(f"[AMBIGUOUS][ATTR] {original_path}")
            for match in matches:
                log_callback(f" -> {match}")
            continue

        stats.missing += 1
        missing_files.append(original_path)
        log_callback(f"[NOT FOUND][ATTR] {original_path}")