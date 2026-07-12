from typing import Callable

APP_NAME = "DCSBPathFixer"

SUPPORTED_AUDIO_EXTENSIONS = {
    ".mp3", ".wav", ".flac", ".ogg", ".oga", ".m4a", ".aac",
    ".wma", ".opus", ".aiff", ".aif", ".ape", ".alac",
    ".mp2", ".mpga", ".m4b",
}

ProgressCallback = Callable[[int, int, str], None]