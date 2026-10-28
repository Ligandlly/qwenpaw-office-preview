# -*- coding: utf-8 -*-
"""DOCX (Word) to HTML rendering.

``mammoth`` is used when available because it understands Word's semantic
styles (headings, lists, tables, footnotes) far better than reading the XML by
hand.  ``python-docx`` provides a conservative fallback so the preview keeps
working in a reduced environment.
"""

from __future__ import annotations

import html
import io
import logging
import re

from .html_doc import document
from .images import image_data_uri
from .sanitize import sanitize_fragment

logger = logging.getLogger(__name__)

# Word documents produced with a localised UI store localised style names.
# Mammoth matches style names, so map the common CJK names explicitly.
_LOCALISED_STYLE_MAP = """
p[style-name='标题 1'] => h1:fresh
p[style-name='标题 2'] => h2:fresh
p[style-name='标题 3'] => h3:fresh
p[style-name='标题 4'] => h4:fresh
p[style-name='标题 5'] => h5:fresh
p[style-name='标题 6'] => h6:fresh
p[style-name='副标题'] => p.qop-subtitle:fresh
p[style-name='引用'] => blockquote:fresh
p[style-name='标题 1 Char'] => h1:fresh
r[style-name='强调'] => em
r[style-name='加粗'] => strong
"""

_MAX_UNSUPPORTED_MESSAGES = 8

# Word stores list numbering either directly on the paragraph or through the
# paragraph style. Mammoth only honours the direct form, so style-driven lists
# ("List Number", "List Bullet", localised equivalents) are promoted here.
_ORDERED_STYLES = {
    "list number",
    "list number 2",
    "list number 3",
    "列表编号",
    "编号列表",
    "listnumber",
}
_BULLET_STYLES = {
    "list bullet",
    "list bullet 2",
    "list bullet 3",
    "list paragraph",
    "列表段落",
    "项目符号",
    "listbullet",
}


class DocxRenderError(RuntimeError):
    """Raised when the DOCX package cannot be parsed at all."""


def _mammoth_convert(data: bytes) -> tuple[str, list[str]]:
    import mammoth
    from mammoth import images

    warnings: list[str] = []

    def _promote_style_numbering(paragraph):
        if paragraph.numbering is not None:
            return paragraph
        name = (paragraph.style_name or "").strip().lower()
        style_id = (paragraph.style_id or "").strip().lower()
        ordered = name in _ORDERED_STYLES or style_id in _ORDERED_STYLES
        bullet = name in _BULLET_STYLES or style_id in _BULLET_STYLES
        if not ordered and not bullet:
            return paragraph
        from mammoth import documents

        return documents.Paragraph(
            children=paragraph.children,
            style_id=paragraph.style_id,
            style_name=paragraph.style_name,
            numbering=documents.numbering_level(0, ordered),
            alignment=paragraph.alignment,
            indent=paragraph.indent,
        )

    def _convert_image(image) -> dict[str, str]:
        try:
            with image.open() as handle:
                blob = handle.read()
        except Exception as exc:  # pragma: no cover - malformed package
            logger.debug("DOCX image could not be read: %s", exc)
            return {"src": ""}
        uri, _, _ = image_data_uri(
            blob,
            getattr(image, "content_type", "image/png"),
        )
        # An empty ``src`` makes browsers request the page itself; use a tiny
        # transparent pixel instead so unsupported vector art disappears.
        if not uri:
            uri = (
                "data:image/gif;base64,"
                "R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw=="
            )
        return {"src": uri}

    styles = """
p[style-name='Title'] => h1:fresh
p[style-name='Subtitle'] => p.qop-subtitle:fresh
p[style-name='Quote'] => blockquote:fresh
p[style-name='Intense Quote'] => blockquote:fresh
""" + _LOCALISED_STYLE_MAP

    with io.BytesIO(data) as stream:
        result = mammoth.convert_to_html(
            stream,
            convert_image=images.img_element(_convert_image),
            style_map=styles,
            transform_document=mammoth.transforms.paragraph(
                _promote_style_numbering,
            ),
        )
    for message in result.messages:
        if getattr(message, "type", "") == "warning":
            warnings.append(str(message.message))
    return result.value, warnings


# ---------------------------------------------------------------------------
# python-docx fallback
# ---------------------------------------------------------------------------

_HEADING_RE = re.compile(r"^(?:heading|标题)\s*([1-9])$", re.IGNORECASE)
_NUMBERING_RE = re.compile(r"^(?:list number|列表编号)", re.IGNORECASE)
_BULLET_RE = re.compile(r"^(?:list bullet|列表段落|项目符号)", re.IGNORECASE)


def _run_html(run, document_part) -> str:
    from docx.oxml.ns import qn

    pieces: list[str] = []
    text = run.text or ""
    if text:
        pieces.append(html.escape(text, quote=False))
    # Inline images live inside the run's XML, not on the run object.
    for blip in run._element.iter(qn("a:blip")):  # noqa: SLF001
        rel_id = blip.get(qn("r:embed"))
        if not rel_id:
            continue
        try:
            part = document_part.related_parts[rel_id]
            uri, _, _ = image_data_uri(part.blob, part.content_type)
        except Exception:  # pragma: no cover - missing relationship
            continue
        if uri:
            alt = html.escape(text or "image", quote=True)
            pieces.append(f'<img src="{uri}" alt="{alt}">')
    if not pieces:
        return ""
    markup = "".join(pieces)
    if run.bold:
        markup = f"<strong>{markup}</strong>"
    if run.italic:
        markup = f"<em>{markup}</em>"
    if run.underline:
        markup = f"<u>{markup}</u>"
    if run.font is not None and run.font.strike:
        markup = f"<s>{markup}</s>"
    return markup


def _paragraph_html(paragraph, document_part) -> tuple[str, str]:
    """Return ``(markup, list_kind)``; ``list_kind`` is ``ul``/``ol``/``""``."""

    body = "".join(_run_html(run, document_part) for run in paragraph.runs)
    if not body:
        # Empty paragraphs are meaningful spacing in Word.
        return "<p>&nbsp;</p>", ""
    style_name = ""
    try:
        style_name = (paragraph.style.name or "").strip()
    except Exception:  # pragma: no cover - broken style table
        style_name = ""
    match = _HEADING_RE.match(style_name)
    if match:
        level = int(match.group(1))
        return f"<h{level}>{body}</h{level}>", ""
    if _NUMBERING_RE.match(style_name):
        return body, "ol"
    if _BULLET_RE.match(style_name):
        return body, "ul"
    lower = style_name.lower()
    if lower == "title":
        return f"<h1>{body}</h1>", ""
    if lower in {"subtitle", "quote", "intense quote"}:
        tag = "blockquote" if "quote" in lower else "p"
        cls = ' class="qop-subtitle"' if lower == "subtitle" else ""
        return f"<{tag}{cls}>{body}</{tag}>", ""
    # Manual numbering / bullets are extremely common in real documents.
    stripped = (paragraph.text or "").lstrip()
    if re.match(r"^[•·▪◦‣∙]\s*", stripped):
        return re.sub(r"^[•·▪◦‣∙]\s*", "", body, count=1), "ul"
    if re.match(r"^\d{1,3}[.)、]\s*", stripped):
        return re.sub(r"^\d{1,3}[.)、]\s*", "", body, count=1), "ol"
    align = ""
    try:
        alignment = paragraph.alignment
        if alignment is not None:
            mapping = {0: "left", 1: "center", 2: "right", 3: "justify"}
            if int(alignment) in mapping:
                align = f' style="text-align:{mapping[int(alignment)]}"'
    except Exception:  # pragma: no cover
        pass
    return f"<p{align}>{body}</p>", ""


def _table_html(table, document_part) -> str:
    rows: list[str] = []
    for row in table.rows:
        cells: list[str] = []
        for cell in row.cells:
            chunks: list[str] = []
            for paragraph in cell.paragraphs:
                markup, _ = _paragraph_html(paragraph, document_part)
                if markup == "<p>&nbsp;</p>":
                    continue
                chunks.append(markup)
            for nested in getattr(cell, "tables", []):
                chunks.append(_table_html(nested, document_part))
            cells.append(f"<td>{''.join(chunks)}</td>")
        rows.append(f"<tr>{''.join(cells)}</tr>")
    return f"<table><tbody>{''.join(rows)}</tbody></table>"


def _python_docx_convert(data: bytes) -> tuple[str, list[str]]:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    try:
        doc = Document(io.BytesIO(data))
    except Exception as exc:
        raise DocxRenderError(f"无法解析 DOCX 文件：{exc}") from exc

    warnings: list[str] = []
    output: list[str] = []
    open_list: str | None = None

    def close_list() -> None:
        nonlocal open_list
        if open_list:
            output.append(f"</{open_list}>")
            open_list = None

    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            paragraph = Paragraph(child, doc)
            # Explicit page breaks split the rendered page.
            if child.findall(f".//{qn('w:br')}"):
                for br in child.iter(qn("w:br")):
                    if br.get(qn("w:type")) == "page":
                        close_list()
                        output.append("<hr>")
            markup, list_kind = _paragraph_html(paragraph, doc.part)
            if list_kind:
                if open_list != list_kind:
                    close_list()
                    output.append(f"<{list_kind}>")
                    open_list = list_kind
                output.append(f"<li>{markup}</li>")
            else:
                close_list()
                output.append(markup)
        elif child.tag == qn("w:tbl"):
            close_list()
            output.append(_table_html(Table(child, doc), doc.part))
    close_list()
    if not any(line.strip() for line in output):
        warnings.append("文档中没有可显示的正文内容。")
    return "".join(output), warnings


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------


def render_docx(
    data: bytes,
    *,
    filename: str,
    theme: str = "auto",
    max_warnings: int = _MAX_UNSUPPORTED_MESSAGES,
) -> tuple[str, list[str]]:
    """Render a DOCX byte string into a complete HTML document.

    Returns ``(html, warnings)``.
    """

    warnings: list[str] = []
    body = ""
    try:
        body, warnings = _mammoth_convert(data)
    except ImportError:
        warnings.append("未安装 mammoth，已改用简化的 Word 渲染器。")
    except DocxRenderError:
        raise
    except Exception as exc:
        logger.info("mammoth failed for %s: %s", filename, exc)
        warnings.append(f"高级 Word 渲染失败（{type(exc).__name__}），已回退到基础渲染。")

    if not body.strip():
        try:
            body, fallback_warnings = _python_docx_convert(data)
        except DocxRenderError:
            raise
        except Exception as exc:
            logger.info("python-docx failed for %s: %s", filename, exc)
            raise DocxRenderError(f"无法读取 Word 文件：{exc}") from exc
        warnings.extend(fallback_warnings)
        body = sanitize_fragment(body)
    else:
        body = sanitize_fragment(body)

    if len(warnings) > max_warnings:
        extra = len(warnings) - max_warnings
        warnings = warnings[:max_warnings] + [f"…另有 {extra} 条格式提示"]

    safe_name = html.escape(filename, quote=False)
    meta = (
        f'<span class="qop-chip">DOCX</span>'
        f'<span class="qop-chip">{safe_name}</span>'
    )
    doc = document(
        f'<main class="qop-doc">{body}</main>',
        title=filename,
        meta=meta,
        warnings=" ".join(html.escape(item, quote=False) for item in warnings),
        theme=theme,
    )
    return doc, warnings
