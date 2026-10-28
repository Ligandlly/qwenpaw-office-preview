# -*- coding: utf-8 -*-
"""PPTX (PowerPoint) to HTML rendering.

Each slide becomes a fixed-size stage (960px wide by default) with absolutely
positioned shapes, so the layout keeps the slide's real proportions.  Text,
tables, pictures, charts (as data tables) and speaker notes are supported.
Everything is plain HTML/CSS - no JavaScript is required inside the preview
iframe.
"""

from __future__ import annotations

import html
import io
import logging
from typing import Any

from .html_doc import pptx_document
from .images import image_data_uri

logger = logging.getLogger(__name__)

TARGET_SLIDE_WIDTH_PX = 960
_DEFAULT_FONT_PT = 18.0
_MAX_TEXT_CHARS = 20_000


class PptxRenderError(RuntimeError):
    """Raised when the PPTX package cannot be parsed."""


def _escape(value: Any) -> str:
    return html.escape(str(value) if value is not None else "", quote=False)


def _emu(value: Any) -> int:
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


def _color_hex(color_format: Any) -> str | None:
    try:
        rgb = color_format.rgb
    except Exception:
        return None
    if rgb is None:
        return None
    return f"#{rgb}"


def _apply_transform(
    shape: Any,
    transform: tuple[float, float, float, float, float, float],
) -> tuple[float, float, float, float]:
    """Map a shape's child-space EMU box into absolute slide EMU space."""

    origin_x, origin_y, scale_x, scale_y, child_x, child_y = transform
    left = (_emu(getattr(shape, "left", 0)) - child_x) * scale_x + origin_x
    top = (_emu(getattr(shape, "top", 0)) - child_y) * scale_y + origin_y
    width = _emu(getattr(shape, "width", 0)) * scale_x
    height = _emu(getattr(shape, "height", 0)) * scale_y
    return left, top, width, height


def _run_html(run: Any, scale: float, default_pt: float = _DEFAULT_FONT_PT) -> str:
    text = run.text or ""
    if not text:
        return ""
    style: list[str] = []
    font = getattr(run, "font", None)
    size_pt = default_pt
    if font is not None:
        try:
            if font.size is not None:
                size_pt = float(font.size.pt)
        except Exception:
            pass
    size_px = max(7.0, min(120.0, size_pt * 12700 * scale))
    style.append(f"font-size:{size_px:.1f}px")
    if font is not None:
        try:
            if font.bold:
                style.append("font-weight:700")
        except Exception:
            pass
        try:
            if font.italic:
                style.append("font-style:italic")
        except Exception:
            pass
        try:
            if font.underline:
                style.append("text-decoration:underline")
        except Exception:
            pass
        try:
            color = _color_hex(font.color)
            if color:
                style.append(f"color:{color}")
        except Exception:
            pass
        try:
            name = str(font.name or "").replace("'", "").replace('"', "")
            if name:
                style.append(f"font-family:'{_escape(name)}',inherit")
        except Exception:
            pass
    markup = _escape(text).replace("\n", "<br>")
    span = f'<span class="qop-run" style="{";".join(style)}">{markup}</span>'
    try:
        address = run.hyperlink.address
    except Exception:
        address = None
    if address and address.lower().startswith(("http://", "https://", "mailto:")):
        span = (
            f'<a href="{html.escape(address, quote=True)}" '
            f'target="_blank" rel="noreferrer">{span}</a>'
        )
    return span


def _default_font_pt(shape: Any) -> float:
    try:
        if getattr(shape, "is_placeholder", False):
            kind = str(shape.placeholder_format.type)
            if "TITLE" in kind:
                return 40.0
            if "SUBTITLE" in kind:
                return 20.0
            if "BODY" in kind:
                return 18.0
    except Exception:
        pass
    return _DEFAULT_FONT_PT


def _text_frame_html(
    text_frame: Any,
    scale: float,
    default_pt: float = _DEFAULT_FONT_PT,
) -> str:
    paragraphs: list[str] = []
    total_chars = 0
    for paragraph in text_frame.paragraphs:
        paragraph_pt = default_pt
        try:
            paragraph_font = paragraph.font
            if paragraph_font is not None and paragraph_font.size is not None:
                paragraph_pt = float(paragraph_font.size.pt)
        except Exception:
            pass
        runs = "".join(
            _run_html(run, scale, paragraph_pt) for run in paragraph.runs
        )
        if not runs and not paragraph.text:
            paragraphs.append("<p>&nbsp;</p>")
            continue
        total_chars += len(paragraph.text)
        if total_chars > _MAX_TEXT_CHARS:
            paragraphs.append(
                '<p class="qop-empty">…文本过长，已截断</p>',
            )
            break
        style: list[str] = ["margin:0"]
        try:
            level = int(paragraph.level or 0)
        except Exception:
            level = 0
        bullet = "• " if level > 0 else ""
        if level > 0:
            style.append(f"padding-left:{level * 18:.0f}px")
        try:
            align = int(paragraph.alignment) if paragraph.alignment is not None else None
        except Exception:
            align = None
        alignment = {1: "left", 2: "center", 3: "right", 4: "justify"}.get(align)
        if alignment:
            style.append(f"text-align:{alignment}")
        paragraphs.append(
            f'<p style="{";".join(style)}">{bullet}{runs}</p>',
        )
    anchor = ""
    try:
        value = str(text_frame.vertical_anchor)
        anchor = {
            "MIDDLE (3)": "center",
            "BOTTOM (4)": "flex-end",
        }.get(value, "flex-start")
    except Exception:
        anchor = "flex-start"
    padding = ""
    try:
        padding = (
            f"padding:{_emu(text_frame.margin_top) * scale:.0f}px "
            f"{_emu(text_frame.margin_right) * scale:.0f}px "
            f"{_emu(text_frame.margin_bottom) * scale:.0f}px "
            f"{_emu(text_frame.margin_left) * scale:.0f}px"
        )
    except Exception:
        padding = "padding:2px 4px"
    return (
        f'<div class="qop-shape-text" style="height:100%;display:flex;'
        f"flex-direction:column;justify-content:{anchor};{padding}\">"
        f"{''.join(paragraphs)}</div>"
    )


def _table_html(table: Any, scale: float) -> str:
    font_px = max(8.0, 12 * 12700 * scale)
    rows: list[str] = []
    for row in table.rows:
        cells: list[str] = []
        for cell in row.cells:
            try:
                if cell.is_spanned:
                    continue
            except Exception:
                pass
            attrs = ""
            try:
                colspan = int(cell.span_width)
                rowspan = int(cell.span_height)
                if colspan > 1:
                    attrs += f' colspan="{colspan}"'
                if rowspan > 1:
                    attrs += f' rowspan="{rowspan}"'
            except Exception:
                pass
            text = _escape(cell.text or "").replace("\n", "<br>")
            cells.append(f"<td{attrs}>{text}</td>")
        rows.append(f"<tr>{''.join(cells)}</tr>")
    if not rows:
        return ""
    return (
        f'<div style="height:100%;overflow:hidden">'
        f'<table style="font-size:{font_px:.1f}px">'
        f"<tbody>{''.join(rows)}</tbody></table></div>"
    )


def _chart_html(chart: Any, scale: float) -> str:
    try:
        categories = [str(item) for item in chart.plots[0].categories]
    except Exception:
        categories = []
    series_rows: list[str] = []
    try:
        for series in chart.series:
            try:
                values = ["" if value is None else value for value in series.values]
            except Exception:
                values = []
            color = ""
            try:
                color_format = series.format.fill.fore_color
                color = _color_hex(color_format) or ""
            except Exception:
                pass
            swatch = (
                f'<span class="qop-swatch" style="background:{color}"></span>'
                if color
                else ""
            )
            name_html = f"{swatch}{_escape(series.name or '')}"
            cells = "".join(f"<td>{_escape(value)}</td>" for value in values)
            series_rows.append(f"<tr><th>{name_html}</th>{cells}</tr>")
    except Exception:
        pass
    if not series_rows and not categories:
        return '<div class="qop-empty">［图表］</div>'
    head = "".join(f"<th>{_escape(name)}</th>" for name in categories)
    return (
        '<div style="height:100%;overflow:auto">'
        f'<table class="qop-chart-table"><thead><tr><th></th>{head}</tr></thead>'
        f"<tbody>{''.join(series_rows)}</tbody></table></div>"
    )


_IDENTITY_TRANSFORM = (0.0, 0.0, 1.0, 1.0, 0.0, 0.0)


def _group_children_html(
    shape: Any,
    scale: float,
    transform: tuple[float, float, float, float, float, float],
) -> str:
    """Render grouped shapes, applying the group's coordinate transform."""

    try:
        children = list(shape.shapes)
    except Exception:
        return ""
    if not children:
        return ""

    left, top, width, height = _apply_transform(shape, transform)

    ch_off_x = ch_off_y = 0.0
    ch_ext_x, ch_ext_y = width, height
    try:
        xfrm = shape._element.grpSpPr.xfrm  # noqa: SLF001
        if xfrm is not None:
            ch_off = xfrm.find(
                "{http://schemas.openxmlformats.org/drawingml/2006/main}chOff",
            )
            ch_ext = xfrm.find(
                "{http://schemas.openxmlformats.org/drawingml/2006/main}chExt",
            )
            if ch_off is not None:
                ch_off_x = float(ch_off.get("x", 0))
                ch_off_y = float(ch_off.get("y", 0))
            if ch_ext is not None:
                ch_ext_x = float(ch_ext.get("cx", width)) or width
                ch_ext_y = float(ch_ext.get("cy", height)) or height
    except Exception:
        pass

    scale_x = transform[2] * (width / ch_ext_x) if ch_ext_x else transform[2]
    scale_y = transform[3] * (height / ch_ext_y) if ch_ext_y else transform[3]
    child_transform = (
        left,
        top,
        scale_x,
        scale_y,
        ch_off_x,
        ch_off_y,
    )
    items: list[str] = []
    for child in children:
        item = _shape_html(child, scale, child_transform)
        if item:
            items.append(item)
    return "".join(items)


def _shape_html(
    shape: Any,
    scale: float,
    transform: tuple[float, float, float, float, float, float] = _IDENTITY_TRANSFORM,
) -> str:
    try:
        shape_type = str(getattr(shape, "shape_type", ""))
    except Exception:
        shape_type = ""

    left, top, width, height = _apply_transform(shape, transform)
    if width <= 0 or height <= 0:
        return ""
    style = (
        f"left:{left * scale:.1f}px;top:{top * scale:.1f}px;"
        f"width:{width * scale:.1f}px;height:{height * scale:.1f}px;"
    )

    # Group
    if "GROUP" in shape_type.upper():
        return _group_children_html(shape, scale, transform)

    # Picture
    try:
        image = getattr(shape, "image", None)
    except Exception:
        image = None
    if image is not None:
        try:
            blob = image.blob
            content_type = image.content_type
            uri, _, _ = image_data_uri(blob, content_type)
        except Exception as exc:
            logger.debug("PPTX image could not be read: %s", exc)
            uri = ""
        if uri:
            alt = _escape(getattr(shape, "name", "image"))
            return (
                f'<div class="qop-shape" style="{style}">'
                f'<img class="qop-shape-img" src="{uri}" alt="{alt}" '
                'style="width:100%;height:100%;object-fit:contain"></div>'
            )
        return ""

    # Table
    if getattr(shape, "has_table", False):
        markup = _table_html(shape.table, scale)
        if markup:
            return f'<div class="qop-shape" style="{style}">{markup}</div>'
        return ""

    # Chart
    if getattr(shape, "has_chart", False):
        markup = _chart_html(shape.chart, scale)
        if markup:
            return f'<div class="qop-shape" style="{style}">{markup}</div>'
        return ""

    # Text (also covers placeholders and autoshapes)
    if getattr(shape, "has_text_frame", False):
        text_frame = shape.text_frame
        if not (text_frame.text or "").strip():
            # Draw the shape outline only when it has a visible fill/line.
            return ""
        return (
            f'<div class="qop-shape" style="{style}">'
            f"{_text_frame_html(text_frame, scale, _default_font_pt(shape))}</div>"
        )

    return ""


def _slide_background(slide: Any) -> str:
    try:
        fill = slide.background.fill
        if int(fill.type) == 1:  # MSO_FILL.SOLID
            color = _color_hex(fill.fore_color)
            if color:
                return color
    except Exception:
        pass
    return ""


def _slide_title(slide: Any) -> str:
    for shape in slide.shapes:
        try:
            if shape.has_text_frame and shape.text_frame.text.strip():
                return shape.text_frame.text.strip().splitlines()[0][:80]
        except Exception:
            continue
    return ""


def render_pptx(
    data: bytes,
    *,
    filename: str,
    theme: str = "auto",
) -> tuple[str, list[str]]:
    from pptx import Presentation

    try:
        presentation = Presentation(io.BytesIO(data))
    except Exception as exc:
        raise PptxRenderError(f"无法解析 PPTX 文件：{exc}") from exc

    try:
        slide_width = _emu(presentation.slide_width) or 9144000
        slide_height = _emu(presentation.slide_height) or 6858000
    except Exception:
        slide_width, slide_height = 9144000, 6858000

    scale = TARGET_SLIDE_WIDTH_PX / slide_width
    stage_height = round(slide_height * scale)

    slides_html: list[str] = []
    notes_html: list[str] = []
    outline: list[str] = []
    warnings: list[str] = []
    unsupported = 0

    for index, slide in enumerate(presentation.slides, start=1):
        title = _slide_title(slide) or f"幻灯片 {index}"
        outline.append(
            f'<li><a href="#slide-{index}">{_escape(title)}</a></li>',
        )
        background = _slide_background(slide)
        bg_style = f"background:{background};" if background else ""
        shapes: list[str] = []
        for shape in slide.shapes:
            try:
                markup = _shape_html(shape, scale)
            except Exception as exc:  # one broken shape must not kill a deck
                logger.debug("PPTX shape failed: %s", exc)
                markup = ""
            if markup:
                shapes.append(markup)
            else:
                try:
                    if not getattr(shape, "has_text_frame", False) or not (
                        shape.text_frame.text or ""
                    ).strip():
                        unsupported += 1
                except Exception:
                    unsupported += 1
        slides_html.append(
            f'<section class="qop-slide" id="slide-{index}" '
            f'style="width:{TARGET_SLIDE_WIDTH_PX}px;height:{stage_height}px;">'
            f'<div class="qop-slide-inner" style="width:{TARGET_SLIDE_WIDTH_PX}px;'
            f'height:{stage_height}px;{bg_style}">{"".join(shapes)}</div>'
            f'<span class="qop-slide-no">{index}</span>'
            "</section>",
        )
        try:
            notes = slide.notes_slide.notes_text_frame.text.strip()
        except Exception:
            notes = ""
        if notes:
            notes_html.append(
                f'<details class="qop-notes" id="notes-{index}">'
                f"<summary>第 {index} 页备注</summary>"
                f"<p>{_escape(notes).replace(chr(10), '<br>')}</p></details>",
            )

    if unsupported:
        warnings.append(f"{unsupported} 个图形元素暂不支持渲染，已跳过。")

    outline_html = (
        f'<nav class="qop-outline"><h2>幻灯片导航（共 {len(slides_html)} 页）</h2>'
        f"<ol>{''.join(outline)}</ol></nav>"
        if slides_html
        else ""
    )
    body = (
        f"{outline_html}<div class=\"qop-slides\">{''.join(slides_html)}</div>"
        f"{''.join(notes_html)}"
    )
    meta = (
        f'<span class="qop-chip">PPTX</span>'
        f'<span class="qop-chip">{_escape(filename)}</span>'
        f'<span class="qop-chip">{len(slides_html)} 页</span>'
    )
    doc = pptx_document(
        body,
        title=filename,
        meta=meta,
        warnings=" ".join(_escape(item) for item in warnings),
        theme=theme,
    )
    return doc, warnings
