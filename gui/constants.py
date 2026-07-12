from core.constants import SUPPORTED_AUDIO_EXTENSIONS

VALID_RENAME_MODES = {"none", "-", "_", "space"}
SUPPORTED_FORMATS_TEXT = ", ".join(sorted(f"*{ext}" for ext in SUPPORTED_AUDIO_EXTENSIONS))