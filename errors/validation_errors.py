from errors.base import DCSBPathFixerError


class ValidationError(DCSBPathFixerError):
    """Raised when user input is invalid."""