# -*- coding: utf-8 -*-
"""Legacy binary Office formats (.doc/.ppt/.xls) via LibreOffice.

Pure-Python libraries for the OLE2 formats are either unmaintained or
incomplete, so legacy files are converted once with a locally installed
LibreOffice (``soffice``) and then rendered by the same modern renderers.
When LibreOffice is unavailable the caller receives a clear, actionable error
instead of a broken preview.
"""

from __future__ import annotations

import logging
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

logger = logging.getLogger(__name__)

LEGACY_TARGETS = {
    ".doc": "docx",
    ".dot": "docx",
    ".wps": "docx",
    ".ppt": "pptx",
    ".pps": "pptx",
    ".dps": "pptx",
    ".xls": "xlsx",
    ".et": "xlsx",
}

_SOFFICE_CANDIDATES = (
    "/usr/bin/libreoffice",
    "/usr/bin/soffice",
    "/usr/local/bin/libreoffice",
    "/usr/local/bin/soffice",
    "/snap/bin/libreoffice",
    "/opt/libreoffice/program/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)


class LegacyConversionError(RuntimeError):
    """Raised when LibreOffice is missing or fails to convert a file."""


def find_libreoffice() -> str | None:
    """Locate a LibreOffice executable, honouring the environment override."""

    override = os.environ.get("QWENPAW_OFFICE_PREVIEW_SOFFICE") or os.environ.get(
        "OFFICE_PREVIEW_SOFFICE",
    )
    if override:
        candidate = Path(override).expanduser()
        if candidate.is_file():
            return str(candidate)
        found = shutil.which(override)
        if found:
            return found
    for name in ("soffice", "libreoffice", "soffice.bin"):
        found = shutil.which(name)
        if found:
            return found
    for candidate in _SOFFICE_CANDIDATES:
        if Path(candidate).is_file():
            return candidate
    return None


def convert_legacy(
    data: bytes,
    *,
    filename: str,
    target: str,
    timeout: float = 180.0,
) -> bytes:
    """Convert *data* to ``docx``/``pptx``/``xlsx`` using LibreOffice."""

    executable = find_libreoffice()
    if not executable:
        raise LegacyConversionError(
            "预览旧版 Office 文件（.doc/.ppt/.xls）需要本机安装 LibreOffice。"
            "请安装 libreoffice 后重试，或在 Office 中另存为 .docx/.pptx/.xlsx。",
        )

    safe_name = Path(filename or f"document.{target}").name
    if not safe_name.lower().endswith(f".{target}"):
        safe_name = f"{Path(safe_name).stem or 'document'}.{target}"
    # Input keeps its original extension so LibreOffice's import filter is
    # chosen correctly; the output name comes from the target format.
    suffix = Path(filename).suffix.lower() or f".{target}"
    source_name = f"source{suffix}"

    with tempfile.TemporaryDirectory(prefix="qwenpaw-office-") as tmp:
        tmp_path = Path(tmp)
        source = tmp_path / source_name
        source.write_bytes(data)
        profile = tmp_path / "lo-profile"
        profile.mkdir()
        command = [
            executable,
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--nodefault",
            "--nofirststartwizard",
            f"-env:UserInstallation={profile.as_uri()}",
            "--convert-to",
            target,
            "--outdir",
            str(tmp_path),
            str(source),
        ]
        env = dict(os.environ)
        env.setdefault("HOME", str(tmp_path))
        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=timeout,
                env=env,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            raise LegacyConversionError(
                f"LibreOffice 转换超时（>{int(timeout)} 秒）。",
            ) from exc
        except OSError as exc:
            raise LegacyConversionError(f"无法启动 LibreOffice：{exc}") from exc

        output = tmp_path / f"{Path(source_name).stem}.{target}"
        if not output.is_file():
            candidates = sorted(tmp_path.glob(f"*.{target}"))
            if candidates:
                output = candidates[0]
        if output.is_file() and output.stat().st_size > 0:
            logger.info(
                "Converted legacy Office file %s to %s via LibreOffice",
                filename,
                target,
            )
            return output.read_bytes()

        stderr = (completed.stderr or "").strip()
        stdout = (completed.stdout or "").strip()
        detail = stderr or stdout or f"exit code {completed.returncode}"
        raise LegacyConversionError(
            f"LibreOffice 转换失败：{detail[:500]}",
        )
