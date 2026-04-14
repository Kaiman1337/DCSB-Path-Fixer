class DCSBPathFixerError(Exception):
    """Base application exception."""


class ValidationError(DCSBPathFixerError):
    """Raised when user input is invalid."""


class ConfigReadError(DCSBPathFixerError):
    """Raised when the XML config cannot be read."""


class ConfigWriteError(DCSBPathFixerError):
    """Raised when the XML config cannot be written."""


class SettingsError(DCSBPathFixerError):
    """Raised when user settings cannot be loaded or saved."""


class HistoryError(DCSBPathFixerError):
    """Raised when history operations fail."""