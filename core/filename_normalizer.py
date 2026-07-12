from __future__ import annotations

import re
from pathlib import Path


def normalize_audio_stem(stem: str, rename_mode: str, use_lowercase: bool = False) -> str:
    original_text = stem.strip()

    cleaned = original_text
    cleaned = re.sub(r"-\(getmp3\.pro\)", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"-\d{5,}", "", cleaned)
    cleaned = cleaned.replace("-+-", "+")
    cleaned = re.sub(r"-{2,}", "-", cleaned)
    cleaned = cleaned.strip(" .-_")
    original_text = cleaned

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
            pass
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

    if use_lowercase:
        text = text.lower()

    return text


def generate_name_variants(filename: str, use_lowercase: bool = False) -> list[str]:
    path = Path(filename)
    stem = path.stem
    suffix = path.suffix

    normalized_space = normalize_audio_stem(stem, "space", use_lowercase)
    normalized_dash = normalize_audio_stem(stem, "-", use_lowercase)
    normalized_underscore = normalize_audio_stem(stem, "_", use_lowercase)
    normalized_none = normalize_audio_stem(stem, "none", use_lowercase)

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