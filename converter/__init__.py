# -*- coding: utf-8 -*-
"""Office document → HTML preview converters for the QwenPaw plugin."""

from .service import (
    DEFAULT_MAX_MB,
    SUPPORTED_EXTENSIONS,
    FileTooLargeError,
    OfficePreviewError,
    UnsupportedFormatError,
    clear_cache,
    convert_office,
    detect_kind,
    health,
    max_bytes,
)

__all__ = [
    "DEFAULT_MAX_MB",
    "SUPPORTED_EXTENSIONS",
    "FileTooLargeError",
    "OfficePreviewError",
    "UnsupportedFormatError",
    "clear_cache",
    "convert_office",
    "detect_kind",
    "health",
    "max_bytes",
]
