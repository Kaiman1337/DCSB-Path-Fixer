from errors.base import DCSBPathFixerError


class ConverterError(DCSBPathFixerError):
    """Raised when audio conversion fails."""


class ConverterDependencyError(ConverterError):
    """Raised when ffmpeg or ffprobe is missing."""