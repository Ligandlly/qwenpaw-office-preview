# -*- coding: utf-8 -*-
"""FastAPI routes exposed at ``/api/office-preview``."""

from __future__ import annotations

import logging

from fastapi import APIRouter, File, HTTPException, Query, Request, UploadFile
from starlette.concurrency import run_in_threadpool

from .service import (
    OfficePreviewError,
    convert_office,
    health,
    max_bytes,
)

logger = logging.getLogger(__name__)

router = APIRouter()


def _error_response(exc: OfficePreviewError) -> HTTPException:
    return HTTPException(
        status_code=exc.status,
        detail={"code": exc.code, "message": str(exc)},
    )


def _filename_from_upload(upload: UploadFile, fallback: str) -> str:
    name = (upload.filename or "").strip()
    return name or fallback


@router.get("/health", summary="Office preview converter status")
def converter_health() -> dict:
    """Report which renderers and optional helpers are available."""

    return health()


@router.post("/convert", summary="Convert an office file to HTML")
async def convert_raw(
    request: Request,
    filename: str = Query(
        default="document.docx",
        max_length=255,
        description="Original file name; the extension selects the renderer.",
    ),
    theme: str = Query(
        default="auto",
        pattern="^(auto|light|dark)$",
        description="Preview colour scheme; defaults to the OS preference.",
    ),
) -> dict:
    """Convert a request body (raw bytes) into a self-contained HTML preview.

    The Console sends the file bytes it already fetched through the
    authenticated workspace/artifact APIs, so the plugin never needs direct
    filesystem access.
    """

    limit = max_bytes()
    content_length = request.headers.get("content-length")
    if content_length:
        try:
            if int(content_length) > limit:
                raise HTTPException(
                    status_code=413,
                    detail={
                        "code": "file_too_large",
                        "message": (
                            f"文件超过 {limit // (1024 * 1024)} MB 上限，无法在线预览。"
                        ),
                    },
                )
        except ValueError:
            pass

    chunks: list[bytes] = []
    total = 0
    async for chunk in request.stream():
        total += len(chunk)
        if total > limit:
            raise HTTPException(
                status_code=413,
                detail={
                    "code": "file_too_large",
                    "message": f"文件超过 {limit // (1024 * 1024)} MB 上限，无法在线预览。",
                },
            )
        chunks.append(chunk)
    data = b"".join(chunks)

    try:
        return await run_in_threadpool(
            convert_office,
            filename,
            data,
            theme=theme,
        )
    except OfficePreviewError as exc:
        raise _error_response(exc) from exc
    except Exception as exc:  # pragma: no cover - defensive
        logger.exception("Unexpected office conversion failure")
        raise HTTPException(
            status_code=500,
            detail={
                "code": "internal_error",
                "message": f"预览生成失败：{exc}",
            },
        ) from exc


@router.post(
    "/convert-upload",
    summary="Convert a multipart office upload to HTML",
)
async def convert_upload(
    file: UploadFile = File(..., description="Office document to preview"),
    theme: str = Query(
        default="auto",
        pattern="^(auto|light|dark)$",
    ),
) -> dict:
    """Multipart variant of :func:`convert_raw` (handy for ``curl``/scripts)."""

    limit = max_bytes()
    chunks: list[bytes] = []
    total = 0
    while chunk := await file.read(1024 * 1024):
        total += len(chunk)
        if total > limit:
            raise HTTPException(
                status_code=413,
                detail={
                    "code": "file_too_large",
                    "message": f"文件超过 {limit // (1024 * 1024)} MB 上限，无法在线预览。",
                },
            )
        chunks.append(chunk)
    data = b"".join(chunks)
    filename = _filename_from_upload(file, "document.docx")

    try:
        return await run_in_threadpool(
            convert_office,
            filename,
            data,
            theme=theme,
        )
    except OfficePreviewError as exc:
        raise _error_response(exc) from exc
    except Exception as exc:  # pragma: no cover - defensive
        logger.exception("Unexpected office conversion failure")
        raise HTTPException(
            status_code=500,
            detail={
                "code": "internal_error",
                "message": f"预览生成失败：{exc}",
            },
        ) from exc
