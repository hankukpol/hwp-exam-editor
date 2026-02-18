class ProcessingError(Exception):
    """Base class for processing failures."""


class UnsupportedFileError(ProcessingError):
    """Raised when a file extension is unsupported."""


class HwpNotAvailableError(ProcessingError):
    """Raised when HWP automation is unavailable on this machine."""


class ParseError(ProcessingError):
    """Raised when parsing fails."""
