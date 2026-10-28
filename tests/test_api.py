# -*- coding: utf-8 -*-
"""HTTP contract tests for the plugin router."""

from __future__ import annotations

import io

import pytest
from docx import Document
from fastapi import FastAPI
from fastapi.testclient import TestClient

from converter.api import router


@pytest.fixture()
def client() -> TestClient:
    app = FastAPI()
    app.include_router(router, prefix="/api/office-preview")
    return TestClient(app)


def docx_bytes() -> bytes:
    doc = Document()
    doc.add_paragraph("Hello QwenPaw")
    stream = io.BytesIO()
    doc.save(stream)
    return stream.getvalue()


def test_health(client: TestClient):
    response = client.get("/api/office-preview/health")
    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is True
    assert ".docx" in payload["supported_extensions"]
    assert "renderers" in payload


def test_convert_raw_bytes(client: TestClient):
    response = client.post(
        "/api/office-preview/convert",
        params={"filename": "hello.docx"},
        content=docx_bytes(),
        headers={"Content-Type": "application/octet-stream"},
    )
    assert response.status_code == 200, response.text
    payload = response.json()
    assert payload["kind"] == "docx"
    assert "Hello QwenPaw" in payload["html"]


def test_convert_multipart(client: TestClient):
    response = client.post(
        "/api/office-preview/convert-upload",
        files={
            "file": (
                "hello.docx",
                docx_bytes(),
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document",
            ),
        },
    )
    assert response.status_code == 200, response.text
    assert "Hello QwenPaw" in response.json()["html"]


def test_unsupported_returns_415(client: TestClient):
    response = client.post(
        "/api/office-preview/convert",
        params={"filename": "image.png"},
        content=b"\x89PNG\r\n",
    )
    assert response.status_code == 415
    detail = response.json()["detail"]
    assert detail["code"] == "unsupported_format"


def test_corrupt_returns_actionable_error(client: TestClient):
    response = client.post(
        "/api/office-preview/convert",
        params={"filename": "broken.docx"},
        content=b"not an office file",
    )
    assert response.status_code in {415, 422}
    assert response.json()["detail"]["message"]


def test_invalid_theme_is_rejected(client: TestClient):
    response = client.post(
        "/api/office-preview/convert",
        params={"filename": "hello.docx", "theme": "neon"},
        content=docx_bytes(),
    )
    assert response.status_code == 422


def test_size_limit_returns_413(client: TestClient, monkeypatch):
    monkeypatch.setenv("QWENPAW_OFFICE_PREVIEW_MAX_MB", "1")
    response = client.post(
        "/api/office-preview/convert",
        params={"filename": "big.docx"},
        content=b"0" * (1024 * 1024 + 1),
    )
    assert response.status_code == 413
    assert response.json()["detail"]["code"] == "file_too_large"
