from errors.base import DCSBPathFixerError


class ConfigReadError(DCSBPathFixerError):
    """Raised when the XML config cannot be read."""


class ConfigWriteError(DCSBPathFixerError):
    """Raised when the XML config cannot be written."""