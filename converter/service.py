# -*- coding: utf-8 -*-
"""Conversion service: sniff the format, pick a renderer, cache the result."""

from __future__ import annotations

import hashlib
import logging
import os
import sys
import threading
import zipfile
from collections import OrderedDict
from io import BytesIO
from pathlib import Path
from typing import Any

from .csv_render import render_csv
from .docx import DocxRenderError, render_docx
from .legacy import (
    LEGACY_TARGETS,
    LegacyConversionError,
    convert_legacy,
    find_libreoffice,
)
from .pptx import PptxRenderError, render_pptx
from .xlsx import XlsxRenderError, render_xlsx

logger = logging.getLogger(__name__)

DEFAULT_MAX_MB = 40
CACHE_ENTRIES = 16

MODERN_KINDS = {
    ".docx": "docx",
    ".docm": "docx",
    ".dotx": "docx",
    ".dotm": "docx",
    ".pptx": "pptx",
    ".pptm": "pptx",
    ".ppsx": "pptx",
    ".potx": "pptx",
    ".xlsx": "xlsx",
    ".xlsm": "xlsx",
    ".xltx": "xlsx",
    ".xltm": "xlsx",
}

CSV_EXTENSIONS = {".csv": ",", ".tsv": "\t"}

SUPPORTED_EXTENSIONS = set(MODERN_KINDS) | set(LEGACY_TARGETS) | set(CSV_EXTENSIONS)

_EXTENSION_LABELS = {
    "docx": "Word",
    "pptx": "PowerPoint",
    "xlsx": "Excel",
    "csv": "CSV",
}


class OfficePreviewError(RuntimeError):
    """Base error carrying a stable machine-readable code and HTTP status."""

    def __init__(self, message: str, *, code: str = "conversion_failed", status: int = 422):
        super().__init__(message)
        self.code = code
        self.status = status


class UnsupportedFormatError(OfficePreviewError):
    def __init__(self, filename: str):
        super().__init__(
            f"不支持预览该文件类型：{filename}。"
            "支持 doc/docx/ppt/pptx/xls/xlsx 及其常见变体，以及 csv/tsv。",
            code="unsupported_format",
            status=415,
        )


class FileTooLargeError(OfficePreviewError):
    def __init__(self, limit_mb: int):
        super().__init__(
            f"文件超过 {limit_mb} MB 上限，无法在线预览。",
            code="file_too_large",
            status=413,
        )


# ---------------------------------------------------------------------------
# Config / health
# ---------------------------------------------------------------------------


def max_bytes() -> int:
    raw = os.environ.get("QWENPAW_OFFICE_PREVIEW_MAX_MB", "")
    try:
        value = int(raw)
    except (TypeError, ValueError):
        value = DEFAULT_MAX_MB
    if value <= 0:
        value = DEFAULT_MAX_MB
    return value * 1024 * 1024


def _module_available(name: str) -> bool:
    """Return True when a renderer dependency can be imported.

    A plain ``find_spec`` probe is not enough: QwenPaw resolves bare imports
    inside a plugin through the plugin's own ``__import__`` hook, so vendored
    packages at the plugin root are importable while ``find_spec`` cannot see
    them.  Importing once is also what the converter will do anyway.
    """

    if name in sys.modules:
        return sys.modules[name] is not None
    try:
        __import__(name)
        return True
    except Exception:
        return False


def health() -> dict[str, Any]:
    soffice = find_libreoffice()
    return {
        "ok": True,
        "supported_extensions": sorted(SUPPORTED_EXTENSIONS),
        "renderers": {
            "docx": _module_available("mammoth") or _module_available("docx"),
            # The pure-Python native renderer is always available.
            "pptx": True,
            "xlsx": _module_available("openpyxl"),
            "csv": True,
            "legacy_doc_ppt_xls": bool(soffice),
        },
        "pptx_engine": (
            "python-pptx" if _module_available("pptx") else "native"
        ),
        "libreoffice": soffice,
        "max_mb": max_bytes() // (1024 * 1024),
        "notes": (
            [
                "未安装 mammoth/python-docx，无法预览 Word 文档。",
            ]
            if not (_module_available("mammoth") or _module_available("docx"))
            else []
        )
        + (
            ["未安装 openpyxl，无法预览 Excel。"]
            if not _module_available("openpyxl")
            else []
        )
        + (
            ["未检测到 LibreOffice，.doc/.ppt/.xls 旧格式无法预览。"]
            if not soffice
            else []
        ),
    }


# ---------------------------------------------------------------------------
# Format sniffing
# ---------------------------------------------------------------------------

_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_ZIP_MAGIC = b"PK\x03\x04"

# Office files are zip containers; a tiny upload must not be allowed to expand
# into hundreds of megabytes of XML. These are deliberately generous: normal
# documents stay far below them.
MAX_UNCOMPRESSED_BYTES = 400 * 1024 * 1024
MAX_COMPRESSION_RATIO = 300


def _guard_archive(data: bytes) -> None:
    try:
        with zipfile.ZipFile(BytesIO(data)) as archive:
            total = 0
            for info in archive.infolist():
                total += info.file_size
                if total > MAX_UNCOMPRESSED_BYTES:
                    raise OfficePreviewError(
                        "文档解压后体积过大，已停止预览。",
                        code="archive_too_large",
                    )
                if (
                    info.compress_size
                    and info.file_size
                    and info.file_size / info.compress_size > 500
                ):
                    raise OfficePreviewError(
                        "文档压缩比异常，已停止预览。",
                        code="archive_too_large",
                    )
            if data and total / len(data) > MAX_COMPRESSION_RATIO:
                raise OfficePreviewError(
                    "文档压缩比异常，已停止预览。",
                    code="archive_too_large",
                )
    except zipfile.BadZipFile:
        return


def _zip_kind(data: bytes) -> str | None:
    try:
        with zipfile.ZipFile(BytesIO(data)) as archive:
            names = set(archive.namelist())
    except (zipfile.BadZipFile, OSError):
        return None
    if "word/document.xml" in names:
        return "docx"
    if "ppt/presentation.xml" in names:
        return "pptx"
    if "xl/workbook.xml" in names:
        return "xlsx"
    return None


def detect_kind(filename: str, data: bytes) -> tuple[str, str]:
    """Return ``(renderer_kind, original_extension)``.

    The byte signature wins over the extension: mislabelled files (a .docx
    saved as .doc, an .xls that is really .xlsx) are common and cheap to
    detect.
    """

    suffix = Path(filename or "").suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        raise UnsupportedFormatError(filename)
    if suffix in CSV_EXTENSIONS:
        return "csv", suffix

    if data.startswith(_ZIP_MAGIC):
        kind = _zip_kind(data)
        if kind:
            return kind, suffix
        # An OOXML container we cannot identify: fall back to the extension.
        modern = MODERN_KINDS.get(suffix)
        if modern:
            return modern, suffix
        raise OfficePreviewError(
            "无法识别该 Office 文件的内部结构，可能已损坏。",
            code="invalid_container",
        )

    if data.startswith(_OLE_MAGIC):
        modern = MODERN_KINDS.get(suffix)
        if modern:
            # A legacy binary container carrying a modern extension.
            return modern, suffix
        target = LEGACY_TARGETS.get(suffix)
        if target:
            return target, suffix
        raise UnsupportedFormatError(filename)

    # Plain-text CSV / TSV.
    if suffix in CSV_EXTENSIONS:
        return "csv", suffix

    raise OfficePreviewError(
        "文件内容不是有效的 Office 文档，可能已损坏或加密。",
        code="invalid_document",
    )


# ---------------------------------------------------------------------------
# Cache
# ---------------------------------------------------------------------------


class _ResultCache:
    def __init__(self, max_entries: int = CACHE_ENTRIES) -> None:
        self._items: OrderedDict[str, dict[str, Any]] = OrderedDict()
        self._lock = threading.Lock()
        self._max = max_entries

    def get(self, key: str) -> dict[str, Any] | None:
        with self._lock:
            item = self._items.get(key)
            if item is not None:
                self._items.move_to_end(key)
            return item

    def put(self, key: str, value: dict[str, Any]) -> None:
        with self._lock:
            self._items[key] = value
            self._items.move_to_end(key)
            while len(self._items) > self._max:
                self._items.popitem(last=False)

    def clear(self) -> None:
        with self._lock:
            self._items.clear()


_cache = _ResultCache()


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------


_REQUIRED_MODULES = {
    "docx": ("mammoth", "docx"),
    "xlsx": ("openpyxl",),
    # PPTX always works: converter/pptx_native.py only needs the stdlib.
}


def _ensure_dependencies(kind: str) -> None:
    """Fail early with an actionable message when renderer deps are missing."""

    modules = _REQUIRED_MODULES.get(kind)
    if not modules:
        return
    if any(_module_available(name) for name in modules):
        return
    raise OfficePreviewError(
        "缺少渲染该文档所需的 Python 依赖"
        f"（{' / '.join(modules)}）。请重新安装或修复插件依赖后重试。",
        code="missing_dependency",
        status=503,
    )


def _render_modern(
    kind: str,
    data: bytes,
    filename: str,
    theme: str,
) -> tuple[str, list[str]]:
    if kind == "docx":
        return render_docx(data, filename=filename, theme=theme)
    if kind == "pptx":
        return render_pptx(data, filename=filename, theme=theme)
    if kind == "xlsx":
        return render_xlsx(data, filename=filename, theme=theme)
    if kind == "csv":
        return render_csv(data, filename=filename, theme=theme)
    raise UnsupportedFormatError(filename)


def _modern_error(kind: str, exc: Exception) -> OfficePreviewError:
    messages = {
        "docx": "Word 文档解析失败",
        "pptx": "PowerPoint 文档解析失败",
        "xlsx": "Excel 工作簿解析失败",
        "csv": "CSV 解析失败",
    }
    wrapping = (DocxRenderError, PptxRenderError, XlsxRenderError)
    if isinstance(exc, wrapping):
        return OfficePreviewError(str(exc), code=f"invalid_{kind}")
    return OfficePreviewError(
        f"{messages.get(kind, '文档')}：{exc}",
        code=f"{kind}_render_failed",
    )


def convert_office(
    filename: str,
    data: bytes,
    *,
    use_cache: bool = True,
    theme: str = "auto",
) -> dict[str, Any]:
    """Convert an office document into a self-contained HTML preview."""

    if not data:
        raise OfficePreviewError("文件为空。", code="empty_file", status=400)
    limit = max_bytes()
    if len(data) > limit:
        raise FileTooLargeError(limit // (1024 * 1024))

    safe_name = Path(filename or "document").name or "document"
    theme = theme if theme in {"light", "dark"} else "auto"
    kind, suffix = detect_kind(safe_name, data)
    if kind in {"docx", "pptx", "xlsx"}:
        _guard_archive(data)

    digest = hashlib.sha256(data).hexdigest()
    cache_key = f"{digest}:{kind}:{theme}:{safe_name}"
    if use_cache:
        cached = _cache.get(cache_key)
        if cached is not None:
            return {**cached, "cached": True}

    warnings: list[str] = []
    converted_from: str | None = None
    render_data = data

    if (
        suffix in LEGACY_TARGETS
        and kind == LEGACY_TARGETS[suffix]
        and data.startswith(_OLE_MAGIC)
    ):
        target = LEGACY_TARGETS[suffix]
        # Mislabeled modern files already returned a modern kind above, so a
        # legacy extension reaching this point really is OLE2.
        try:
            render_data = convert_legacy(
                data,
                filename=safe_name,
                target=target,
            )
        except LegacyConversionError as exc:
            raise OfficePreviewError(
                str(exc),
                code="legacy_converter_missing",
                status=501,
            ) from exc
        converted_from = suffix
        kind = target
        warnings.append(
            f"已通过 LibreOffice 将 {suffix} 转换为 {target} 后渲染。",
        )

    _ensure_dependencies(kind)

    try:
        html, render_warnings = _render_modern(
            kind,
            render_data,
            safe_name,
            theme,
        )
    except OfficePreviewError:
        raise
    except Exception as exc:
        logger.info("Native render failed for %s: %s", safe_name, exc)
        soffice = find_libreoffice()
        if soffice and kind in {"docx", "pptx", "xlsx"} and converted_from is None:
            # A damaged OOXML package sometimes opens fine after a round-trip
            # through LibreOffice; try that once before surfacing the error.
            try:
                repaired = convert_legacy(
                    data,
                    filename=safe_name,
                    target=kind,
                )
                html, render_warnings = _render_modern(
                    kind,
                    repaired,
                    safe_name,
                    theme,
                )
                warnings.append("文档结构异常，已通过 LibreOffice 修复后渲染。")
            except Exception:
                raise _modern_error(kind, exc) from exc
        else:
            raise _modern_error(kind, exc) from exc
    else:
        render_warnings = render_warnings or []

    warnings.extend(item for item in render_warnings if item not in warnings)

    meta = {
        "kind": kind,
        "label": _EXTENSION_LABELS.get(kind, kind.upper()),
        "extension": suffix.lstrip("."),
        "converted_from": converted_from,
        "size": len(data),
        "html_size": len(html),
    }
    result = {
        "ok": True,
        "filename": safe_name,
        "kind": kind,
        "format": "html",
        "html": html,
        "warnings": warnings,
        "meta": meta,
        "cached": False,
    }
    # Very large previews (image-heavy decks) would make the in-memory cache
    # expensive; only cache results of a reasonable size.
    if use_cache and len(html) <= 4 * 1024 * 1024:
        _cache.put(cache_key, result)
    return result


def clear_cache() -> None:
    _cache.clear()
