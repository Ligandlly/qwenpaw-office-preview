# -*- coding: utf-8 -*-
"""Image helpers shared by the DOCX / PPTX renderers."""

from __future__ import annotations

import base64
import io
import logging

logger = logging.getLogger(__name__)

# Keeping the embedded payload small matters: the whole preview travels as a
# single JSON response and is injected into the page as inline markup.
DEFAULT_MAX_IMAGE_PX = 1600
DEFAULT_JPEG_QUALITY = 82


def _looks_like_vector(blob: bytes) -> bool:
    head = blob[:8]
    # EMF (0x01 0x00 0x00 0x00) / WMF (0xD7 0xCD 0xC6 0x9A) are not
    # universally supported by Pillow; browsers cannot render them either.
    return head.startswith(b"\xd7\xcd\xc6\x9a") or head[:4] == b"\x01\x00\x00\x00"


def image_data_uri(
    blob: bytes,
    content_type: str = "image/png",
    *,
    max_px: int = DEFAULT_MAX_IMAGE_PX,
) -> tuple[str, int | None, int | None]:
    """Return ``(data_uri, width, height)`` for *blob*.

    Raster images are downscaled to *max_px* when Pillow is available.  Vector
    or unsupported formats are embedded as-is; callers can decide to skip them
    when a browser cannot display them.
    """

    content_type = (content_type or "image/png").split(";")[0].strip().lower()
    if _looks_like_vector(blob):
        return "", None, None
    try:
        from PIL import Image
    except Exception:  # pragma: no cover - Pillow is a declared dependency
        Image = None  # type: ignore[assignment]

    if Image is not None:
        try:
            with Image.open(io.BytesIO(blob)) as probe:
                width, height = probe.size
                has_alpha = probe.mode in {"RGBA", "LA"} or (
                    probe.mode == "P" and "transparency" in probe.info
                )
                if max(width, height) > max_px:
                    ratio = max_px / float(max(width, height))
                    size = (max(1, int(width * ratio)), max(1, int(height * ratio)))
                    probe.load()
                    resized = probe.resize(size, Image.LANCZOS)
                    out = io.BytesIO()
                    if has_alpha:
                        resized.save(out, format="PNG", optimize=True)
                        content_type = "image/png"
                    else:
                        resized.convert("RGB").save(
                            out,
                            format="JPEG",
                            quality=DEFAULT_JPEG_QUALITY,
                            optimize=True,
                        )
                        content_type = "image/jpeg"
                    blob = out.getvalue()
                    width, height = size
                else:
                    # Re-encode oversized originals (e.g. multi-megabyte PNG
                    # screenshots) while keeping transparency.
                    if len(blob) > 1_200_000:
                        probe.load()
                        out = io.BytesIO()
                        if has_alpha:
                            probe.save(out, format="PNG", optimize=True)
                            content_type = "image/png"
                        else:
                            probe.convert("RGB").save(
                                out,
                                format="JPEG",
                                quality=DEFAULT_JPEG_QUALITY,
                                optimize=True,
                            )
                            content_type = "image/jpeg"
                        candidate = out.getvalue()
                        if len(candidate) < len(blob):
                            blob = candidate
        except Exception as exc:  # unsupported format or corrupt image
            logger.debug("Could not normalise image (%s): %s", content_type, exc)
            if content_type in {"image/x-emf", "image/x-wmf", "image/emf", "image/wmf"}:
                return "", None, None
            width = height = None  # type: ignore[assignment]

    encoded = base64.b64encode(blob).decode("ascii")
    return f"data:{content_type};base64,{encoded}", width, height
