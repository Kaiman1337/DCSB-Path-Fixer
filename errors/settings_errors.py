from errors.base import DCSBPathFixerError


class SettingsError(DCSBPathFixerError):
    """Raised when user settings cannot be loaded or saved."""