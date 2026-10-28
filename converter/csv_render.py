# -*- coding: utf-8 -*-
"""CSV / TSV rendering (bonus format, reuses the spreadsheet look)."""

from __future__ import annotations

import csv
import html
import io
from typing import Any

from .html_doc import xlsx_document

MAX_ROWS = 500
MAX_COLS = 60


def _escape(value: Any) -> str:
    return html.escape(str(value) if value is not None else "", quote=False)


def _column_label(index: int) -> str:
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    label = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        label = letters[rem] + label
    return label


def _decode(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "gb18030", "utf-16", "latin-1"):
        try:
            return data.decode(encoding)
        except (UnicodeDecodeError, LookupError):
            continue
    return data.decode("utf-8", errors="replace")


def render_csv(
    data: bytes,
    *,
    filename: str,
    theme: str = "auto",
) -> tuple[str, list[str]]:
    text = _decode(data)
    sample = text[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        delimiter = dialect.delimiter
    except csv.Error:
        delimiter = "\t" if filename.lower().endswith(".tsv") else ","

    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    rows: list[list[str]] = []
    truncated = False
    for index, row in enumerate(reader):
        if index >= MAX_ROWS:
            truncated = True
            break
        rows.append(row[:MAX_COLS])

    warnings: list[str] = []
    if truncated:
        warnings.append(f"仅显示前 {MAX_ROWS} 行。")
    if any(len(row) > MAX_COLS for row in rows):
        warnings.append(f"仅显示前 {MAX_COLS} 列。")

    if not rows:
        rows = [[""]]
    header, body = rows[0], rows[1:]
    col_count = max(len(row) for row in rows) or 1
    header_html = "".join(
        f"<th>{_escape(header[col]) if col < len(header) else ''}</th>"
        for col in range(col_count)
    )
    body_html: list[str] = []
    for row_index, row in enumerate(body, start=1):
        cells = "".join(
            f"<td>{_escape(row[col]) if col < len(row) else ''}</td>"
            for col in range(col_count)
        )
        body_html.append(f"<tr><th scope=\"row\">{row_index}</th>{cells}</tr>")

    table = (
        '<section class="qop-sheet"><h2>'
        f"{_escape(filename)}</h2>"
        '<div class="qop-sheet-scroll"><table class="qop-grid">'
        '<colgroup><col style="width:52px"></colgroup>'
        '<thead><tr><th class="qop-corner">#</th>'
        f"{header_html}</tr></thead>"
        f"<tbody>{''.join(body_html)}</tbody></table></div></section>"
    )
    meta = (
        f'<span class="qop-chip">CSV</span>'
        f'<span class="qop-chip">{_escape(filename)}</span>'
        f'<span class="qop-chip">{len(body)} 行 × {col_count} 列</span>'
    )
    doc = xlsx_document(
        f'<div class="qop-workbook">{table}</div>',
        title=filename,
        meta=meta,
        warnings=" ".join(_escape(item) for item in warnings),
        theme=theme,
    )
    return doc, warnings
