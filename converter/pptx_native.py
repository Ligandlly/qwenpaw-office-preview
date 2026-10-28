# -*- coding: utf-8 -*-
"""Pure-Python PPTX renderer used when ``python-pptx`` is unavailable.

The offline plugin bundle intentionally vendors only pure-Python packages, so
the usual ``python-pptx`` + ``lxml`` stack cannot be shipped.  This module
reads the OOXML package with :mod:`zipfile` / :mod:`xml.etree` and produces the
same slide markup as :mod:`converter.pptx`.

Supported: slide order/size, text boxes and placeholders (runs, colours, bold/
italic/underline, alignment, levels, bullets), solid fills, preset shapes,
group shapes, pictures, tables, simple charts as data tables, slide background
and speaker notes.  Animations, SmartArt, embedded OLE objects and transitions
are out of scope.
"""

from __future__ import annotations

import html
import io
import logging
import posixpath
import re
import zipfile
from dataclasses import dataclass
from typing import Any
from xml.etree import ElementTree as ET

from .html_doc import pptx_document
from .images import image_data_uri

logger = logging.getLogger(__name__)

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

TARGET_SLIDE_WIDTH_PX = 960
_DEFAULT_FONT_PT = 18.0


def _cqw(px: float) -> str:
    return f"{px / TARGET_SLIDE_WIDTH_PX * 100:.3f}cqw"


def _percent(px: float, total: float) -> str:
    if not total:
        return "0%"
    return f"{px / total * 100:.3f}%"
_MAX_SLIDES = 300
_MAX_CHART_POINTS = 200

_THEME_DEFAULTS = {
    "dk1": "#000000",
    "lt1": "#FFFFFF",
    "dk2": "#44546A",
    "lt2": "#E7E6E6",
    "accent1": "#4472C4",
    "accent2": "#ED7D31",
    "accent3": "#A5A5A5",
    "accent4": "#FFC000",
    "accent5": "#5B9BD5",
    "accent6": "#70AD47",
    "hlink": "#0563C1",
    "folHlink": "#954F72",
    "tx1": "#000000",
    "tx2": "#44546A",
    "bg1": "#FFFFFF",
    "bg2": "#E7E6E6",
}


class PptxNativeError(RuntimeError):
    """Raised when the PPTX package cannot be parsed."""


def _q(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def _escape(value: Any) -> str:
    return html.escape(str(value) if value is not None else "", quote=False)


def _int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _rels_path(part: str) -> str:
    directory, name = posixpath.split(part)
    return posixpath.join(directory, "_rels", f"{name}.rels")


def _resolve_part(base_part: str, target: str) -> str:
    target = (target or "").replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(
        posixpath.join(posixpath.dirname(base_part), target),
    )


@dataclass(frozen=True)
class _Rel:
    id: str
    type: str
    target: str
    external: bool


class _Package:
    def __init__(self, data: bytes) -> None:
        try:
            self._zip = zipfile.ZipFile(io.BytesIO(data))
        except (zipfile.BadZipFile, OSError) as exc:
            raise PptxNativeError(f"无法打开 PPTX 压缩包：{exc}") from exc
        self.names = set(self._zip.namelist())
        self._xml: dict[str, ET.Element] = {}
        self._rels: dict[str, dict[str, _Rel]] = {}
        self._overrides, self._defaults = self._load_content_types()

    def close(self) -> None:
        try:
            self._zip.close()
        except Exception:  # pragma: no cover
            pass

    def read(self, part: str) -> bytes:
        return self._zip.read(part)

    def xml(self, part: str) -> ET.Element:
        if part not in self._xml:
            try:
                self._xml[part] = ET.fromstring(self._zip.read(part))
            except (KeyError, ET.ParseError) as exc:
                raise PptxNativeError(
                    f"无法解析 PPTX 部件 {part}：{exc}",
                ) from exc
        return self._xml[part]

    def _load_content_types(self) -> tuple[dict[str, str], dict[str, str]]:
        overrides: dict[str, str] = {}
        defaults: dict[str, str] = {}
        try:
            root = self.xml("[Content_Types].xml")
        except PptxNativeError:
            return overrides, defaults
        for element in root:
            if element.tag == _q("ct", "Override"):
                part = (element.get("PartName") or "").lstrip("/")
                if part:
                    overrides[part] = element.get("ContentType") or ""
            elif element.tag == _q("ct", "Default"):
                extension = (element.get("Extension") or "").lower()
                if extension:
                    defaults[extension] = element.get("ContentType") or ""
        return overrides, defaults

    def content_type(self, part: str) -> str:
        if part in self._overrides:
            return self._overrides[part]
        extension = part.rsplit(".", 1)[-1].lower() if "." in part else ""
        return self._defaults.get(extension, "application/octet-stream")

    def rels(self, part: str) -> dict[str, _Rel]:
        if part in self._rels:
            return self._rels[part]
        relations: dict[str, _Rel] = {}
        rels_part = _rels_path(part)
        if rels_part in self.names:
            try:
                root = self.xml(rels_part)
            except PptxNativeError:
                root = None
            if root is not None:
                for element in root:
                    if element.tag != _q("rel", "Relationship"):
                        continue
                    rel_id = element.get("Id") or ""
                    if not rel_id:
                        continue
                    target = element.get("Target") or ""
                    relations[rel_id] = _Rel(
                        id=rel_id,
                        type=element.get("Type") or "",
                        target=_resolve_part(part, target),
                        external=(element.get("TargetMode") or "").lower()
                        == "external",
                    )
        self._rels[part] = relations
        return relations


# ---------------------------------------------------------------------------
# Theme / colours
# ---------------------------------------------------------------------------


def _load_theme(package: _Package) -> tuple[dict[str, str], dict[str, str]]:
    colors = dict(_THEME_DEFAULTS)
    fonts = {"major": "", "minor": ""}
    presentation = "ppt/presentation.xml"
    theme_part = ""
    for rel in package.rels(presentation).values():
        if rel.type.endswith("/theme") and not rel.external:
            theme_part = rel.target
            break
    if not theme_part or theme_part not in package.names:
        theme_part = (
            "ppt/theme/theme1.xml" if "ppt/theme/theme1.xml" in package.names else ""
        )
    if not theme_part:
        return colors, fonts
    try:
        root = package.xml(theme_part)
    except PptxNativeError:
        return colors, fonts
    scheme = root.find(f"./{_q('a', 'themeElements')}/{_q('a', 'clrScheme')}")
    if scheme is not None:
        for entry in scheme:
            name = entry.tag.split("}", 1)[-1]
            for child in entry:
                value = child.get("val") or child.get("lastClr") or ""
                if child.tag == _q("a", "srgbClr") and value:
                    colors[name] = f"#{value}"
                elif child.tag == _q("a", "sysClr") and value:
                    colors[name] = f"#{value}"
    font_scheme = root.find(f"./{_q('a', 'themeElements')}/{_q('a', 'fontScheme')}")
    if font_scheme is not None:
        for key, tag in (("major", "majorFont"), ("minor", "minorFont")):
            node = font_scheme.find(_q("a", tag))
            if node is None:
                continue
            latin = node.find(_q("a", "latin"))
            if latin is not None:
                fonts[key] = latin.get("typeface") or ""
    return colors, fonts


def _color_from(container: ET.Element | None, theme: dict[str, str]) -> str | None:
    if container is None:
        return None
    for child in container:
        tag = child.tag
        if tag == _q("a", "srgbClr"):
            value = (child.get("val") or "").strip()
            return f"#{value}" if re.fullmatch(r"[0-9a-fA-F]{6,8}", value) else None
        if tag == _q("a", "schemeClr"):
            value = (child.get("val") or "").strip()
            return theme.get(value)
        if tag == _q("a", "sysClr"):
            value = (child.get("val") or "").strip()
            if value:
                return theme.get(value)
            last = (child.get("lastClr") or "").strip()
            return f"#{last}" if re.fullmatch(r"[0-9a-fA-F]{6}", last) else None
        if tag == _q("a", "prstClr"):
            value = (child.get("val") or "").strip()
            return value if re.fullmatch(r"[a-zA-Z]{3,20}", value) else None
    return None


def _fill_color(element: ET.Element | None, theme: dict[str, str]) -> str | None:
    if element is None:
        return None
    solid = element.find(_q("a", "solidFill"))
    return _color_from(solid, theme)


# ---------------------------------------------------------------------------
# Geometry
# ---------------------------------------------------------------------------

_TRANSFORM = tuple[float, float, float, float, float, float]


def _xfrm_box(
    xfrm: ET.Element | None,
    transform: _TRANSFORM,
) -> tuple[float, float, float, float]:
    if xfrm is None:
        return 0.0, 0.0, 0.0, 0.0
    off = xfrm.find(_q("a", "off"))
    ext = xfrm.find(_q("a", "ext"))
    origin_x, origin_y, scale_x, scale_y, child_x, child_y = transform
    left = ((_int(off.get("x")) if off is not None else 0) - child_x) * scale_x
    top = ((_int(off.get("y")) if off is not None else 0) - child_y) * scale_y
    width = (_int(ext.get("cx")) if ext is not None else 0) * scale_x
    height = (_int(ext.get("cy")) if ext is not None else 0) * scale_y
    return left + origin_x, top + origin_y, width, height


def _rotation(xfrm: ET.Element | None) -> str:
    if xfrm is None:
        return ""
    rot = _int(xfrm.get("rot"), 0) / 60000.0
    flip_h = xfrm.get("flipH") in {"1", "true"}
    flip_v = xfrm.get("flipV") in {"1", "true"}
    pieces: list[str] = []
    if rot:
        pieces.append(f"rotate({rot:.2f}deg)")
    if flip_h:
        pieces.append("scaleX(-1)")
    if flip_v:
        pieces.append("scaleY(-1)")
    return ";".join(pieces)


def _shape_style(
    box: tuple[float, float, float, float],
    scale: float,
    stage_height: float,
    *,
    background: str | None = None,
    border: str | None = None,
    transform: str = "",
) -> str:
    left, top, width, height = box
    declarations = [
        f"left:{_percent(left * scale, TARGET_SLIDE_WIDTH_PX)}",
        f"top:{_percent(top * scale, stage_height)}",
        f"width:{_percent(width * scale, TARGET_SLIDE_WIDTH_PX)}",
        f"height:{_percent(height * scale, stage_height)}",
    ]
    if background:
        declarations.append(f"background:{background}")
    if border:
        declarations.append(f"border:{border}")
    if transform:
        declarations.append(f"transform:{transform}")
        declarations.append("transform-origin:center center")
    return ";".join(declarations)


# ---------------------------------------------------------------------------
# Text
# ---------------------------------------------------------------------------

_ALIGNMENT = {"l": "left", "ctr": "center", "r": "right", "just": "justify"}


def _placeholder_type(element: ET.Element | None) -> str:
    if element is None:
        return ""
    return (element.get("type") or "body").strip()


def _default_font_pt(placeholder: str) -> float:
    if placeholder in {"title", "ctrTitle"}:
        return 40.0
    if placeholder == "subTitle":
        return 20.0
    return _DEFAULT_FONT_PT


@dataclass
class _TextContext:
    theme: dict[str, str]
    fonts: dict[str, str]
    scale: float
    rels: dict[str, _Rel]
    placeholder: str = ""
    list_counters: dict[int, int] | None = None

    def __post_init__(self) -> None:
        if self.list_counters is None:
            self.list_counters = {}


def _font_family(element: ET.Element | None, ctx: _TextContext, major: bool) -> str:
    if element is not None:
        latin = element.find(_q("a", "latin"))
        if latin is not None:
            typeface = (latin.get("typeface") or "").strip()
            if typeface and not typeface.startswith("+"):
                return typeface.replace("'", "").replace('"', "")
    typeface = ctx.fonts.get("major" if major else "minor", "")
    return typeface.replace("'", "").replace('"', "")


def _run_html(run: ET.Element, ctx: _TextContext, default_pt: float) -> str:
    rpr = run.find(_q("a", "rPr"))
    text = "".join(node.text or "" for node in run.findall(_q("a", "t")))
    if not text:
        return ""
    size_pt = default_pt
    if rpr is not None and rpr.get("sz"):
        size_pt = _int(rpr.get("sz"), int(default_pt * 100)) / 100.0
    size_px = max(7.0, min(120.0, size_pt * 12700 * ctx.scale))
    declarations = [f"font-size:{_cqw(size_px)}"]
    if rpr is not None:
        if rpr.get("b") in {"1", "true"}:
            declarations.append("font-weight:700")
        if rpr.get("i") in {"1", "true"}:
            declarations.append("font-style:italic")
        decoration: list[str] = []
        if rpr.get("u") and rpr.get("u") != "none":
            decoration.append("underline")
        if rpr.get("strike") and rpr.get("strike") != "noStrike":
            decoration.append("line-through")
        if decoration:
            declarations.append(f"text-decoration:{' '.join(decoration)}")
        color = _fill_color(rpr, ctx.theme)
        if color:
            declarations.append(f"color:{color}")
        family = _font_family(rpr, ctx, major=ctx.placeholder in {"title", "ctrTitle"})
        if family:
            declarations.append(f"font-family:'{_escape(family)}',inherit")
    markup = _escape(text).replace("\n", "<br>")
    span = f'<span class="qop-run" style="{";".join(declarations)}">{markup}</span>'
    hlink = rpr.find(_q("a", "hlinkClick")) if rpr is not None else None
    rel_id = hlink.get(_q("r", "id")) if hlink is not None else None
    rel = ctx.rels.get(rel_id or "")
    if rel is not None and rel.external and rel.target.lower().startswith(
        ("http://", "https://", "mailto:"),
    ):
        span = (
            f'<a href="{html.escape(rel.target, quote=True)}" '
            f'target="_blank" rel="noreferrer">{span}</a>'
        )
    return span


def _paragraph_html(
    paragraph: ET.Element,
    ctx: _TextContext,
    default_pt: float,
) -> str:
    ppr = paragraph.find(_q("a", "pPr"))
    level = _int(ppr.get("lvl")) if ppr is not None else 0
    default_run = (
        ppr.find(_q("a", "defRPr")) if ppr is not None else None
    )
    run_default_pt = default_pt
    if default_run is not None and default_run.get("sz"):
        run_default_pt = _int(default_run.get("sz"), int(default_pt * 100)) / 100.0

    pieces: list[str] = []
    bullet = ""
    if ppr is not None:
        if ppr.find(_q("a", "buNone")) is not None:
            bullet = ""
        elif ppr.find(_q("a", "buAutoNum")) is not None:
            counters = ctx.list_counters or {}
            counters[level] = counters.get(level, 0) + 1
            bullet = f"{counters[level]}. "
        else:
            bu_char = ppr.find(_q("a", "buChar"))
            if bu_char is not None:
                bullet = (bu_char.get("char") or "•") + " "
            elif level > 0:
                bullet = "• "
    if bullet:
        pieces.append(f'<span class="qop-run">{_escape(bullet)}</span>')

    for child in paragraph:
        if child.tag == _q("a", "r"):
            pieces.append(_run_html(child, ctx, run_default_pt))
        elif child.tag == _q("a", "br"):
            pieces.append("<br>")
        elif child.tag == _q("a", "fld"):
            text = "".join(node.text or "" for node in child.findall(_q("a", "t")))
            if text:
                pieces.append(_escape(text))

    if not pieces:
        pieces.append("&nbsp;")
    style = ["margin:0"]
    if level > 0:
        style.append(f"padding-left:{_cqw(level * 18)}")
    if ppr is not None:
        align = ppr.get("algn")
        if align in _ALIGNMENT:
            style.append(f"text-align:{_ALIGNMENT[align]}")
        spacing = ppr.find(f"{_q('a', 'lnSpc')}/{_q('a', 'spcPct')}")
        if spacing is not None:
            percent = _int(spacing.get("val"), 100000) / 1000.0
            if 50 <= percent <= 300:
                style.append(f"line-height:{percent / 100:.2f}")
        space_before = ppr.find(f"{_q('a', 'spcBef')}/{_q('a', 'spcPts')}")
        if space_before is not None:
            points = _int(space_before.get("val"), 0) / 100.0
            if points:
                style.append(f"margin-top:{points:.1f}pt")
    return f'<p style="{";".join(style)}">{"".join(pieces)}</p>'


def _text_frame_html(
    tx_body: ET.Element | None,
    ctx: _TextContext,
    *,
    default_pt: float | None = None,
) -> str:
    if tx_body is None:
        return ""
    body_pr = tx_body.find(_q("a", "bodyPr"))
    anchor = (body_pr.get("anchor") if body_pr is not None else "") or "t"
    justify = {"ctr": "center", "b": "flex-end"}.get(anchor, "flex-start")
    margins = {"lIns": 91440, "rIns": 91440, "tIns": 45720, "bIns": 45720}
    if body_pr is not None:
        for key in margins:
            if body_pr.get(key) is not None:
                margins[key] = _int(body_pr.get(key), margins[key])
    padding = (
        f"{_cqw(margins['tIns'] * ctx.scale)} "
        f"{_cqw(margins['rIns'] * ctx.scale)} "
        f"{_cqw(margins['bIns'] * ctx.scale)} "
        f"{_cqw(margins['lIns'] * ctx.scale)}"
    )
    base_pt = default_pt if default_pt is not None else _default_font_pt(
        ctx.placeholder,
    )
    paragraphs = [
        _paragraph_html(paragraph, ctx, base_pt)
        for paragraph in tx_body.findall(_q("a", "p"))
    ]
    if not paragraphs:
        return ""
    return (
        f'<div class="qop-shape-text" style="height:100%;display:flex;'
        f"flex-direction:column;justify-content:{justify};padding:{padding}\">"
        f"{''.join(paragraphs)}</div>"
    )


# ---------------------------------------------------------------------------
# Tables / charts
# ---------------------------------------------------------------------------


def _table_html(
    table: ET.Element,
    ctx: _TextContext,
    box: tuple[float, float, float, float],
) -> str:
    rows_html: list[str] = []
    for row in table.findall(_q("a", "tr")):
        cells_html: list[str] = []
        for cell in row.findall(_q("a", "tc")):
            if cell.get("hMerge") in {"1", "true"} or cell.get("vMerge") in {
                "1",
                "true",
            }:
                continue
            attrs = ""
            colspan = _int(cell.get("gridSpan"), 1)
            rowspan = _int(cell.get("rowSpan"), 1)
            if colspan > 1:
                attrs += f' colspan="{colspan}"'
            if rowspan > 1:
                attrs += f' rowspan="{rowspan}"'
            cell_pr = cell.find(_q("a", "tcPr"))
            fill = _fill_color(cell_pr, ctx.theme)
            if fill:
                attrs += f' style="background:{fill}"'
            cell_ctx = _TextContext(
                theme=ctx.theme,
                fonts=ctx.fonts,
                scale=ctx.scale,
                rels=ctx.rels,
                placeholder="",
            )
            text = _text_frame_html(
                cell.find(_q("a", "txBody")),
                cell_ctx,
                default_pt=12.0,
            )
            cells_html.append(f"<td{attrs}>{text}</td>")
        rows_html.append(f"<tr>{''.join(cells_html)}</tr>")
    if not rows_html:
        return ""
    return (
        f'<table style="width:100%;height:100%;border-collapse:collapse">'
        f"<tbody>{''.join(rows_html)}</tbody></table>"
    )


def _chart_html(
    package: _Package,
    rel: _Rel,
    ctx: _TextContext,
) -> str:
    if rel.external or rel.target not in package.names:
        return '<div class="qop-empty">［图表］</div>'
    try:
        root = package.xml(rel.target)
    except PptxNativeError:
        return '<div class="qop-empty">［图表］</div>'
    categories: list[str] = []
    try:
        cat = root.find(f".//{_q('c', 'cat')}")
        if cat is not None:
            for point in cat.findall(f".//{_q('c', 'pt')}"):
                categories.append(point.findtext(_q("c", "v")) or "")
    except Exception:
        pass
    series_html: list[str] = []
    try:
        for series in root.findall(f".//{_q('c', 'ser')}"):
            name = ""
            tx = series.find(_q("c", "tx"))
            if tx is not None:
                name = tx.findtext(f".//{_q('c', 'v')}") or ""
            values: list[str] = []
            val = series.find(_q("c", "val"))
            if val is not None:
                for point in val.findall(f".//{_q('c', 'pt')}"):
                    values.append(point.findtext(_q("c", "v")) or "")
            color = _fill_color(series.find(_q("c", "spPr")), ctx.theme)
            swatch = (
                f'<span class="qop-swatch" style="background:{color}"></span>'
                if color
                else ""
            )
            cells = "".join(
                f"<td>{_escape(value)}</td>"
                for value in values[:_MAX_CHART_POINTS]
            )
            series_html.append(f"<tr><th>{swatch}{_escape(name)}</th>{cells}</tr>")
    except Exception:
        pass
    if not series_html and not categories:
        return '<div class="qop-empty">［图表］</div>'
    head = "".join(f"<th>{_escape(name)}</th>" for name in categories)
    return (
        '<div style="height:100%;overflow:auto">'
        f'<table class="qop-chart-table"><thead><tr><th></th>{head}</tr></thead>'
        f"<tbody>{''.join(series_html)}</tbody></table></div>"
    )


# ---------------------------------------------------------------------------
# Shapes
# ---------------------------------------------------------------------------


def _placeholder_node(shape: ET.Element) -> ET.Element | None:
    for path in (
        f"./{_q('p', 'nvSpPr')}/{_q('p', 'nvPr')}/{_q('p', 'ph')}",
        f"./{_q('p', 'nvPicPr')}/{_q('p', 'nvPr')}/{_q('p', 'ph')}",
        f"./{_q('p', 'nvGraphicFramePr')}/{_q('p', 'nvPr')}/{_q('p', 'ph')}",
    ):
        node = shape.find(path)
        if node is not None:
            return node
    return None


def _placeholder_of(shape: ET.Element) -> str:
    node = _placeholder_node(shape)
    return _placeholder_type(node) if node is not None else ""


def _placeholder_key(shape: ET.Element) -> tuple[str, str] | None:
    node = _placeholder_node(shape)
    if node is None:
        return None
    return (_placeholder_type(node), (node.get("idx") or "").strip())


def _placeholder_xfrms(
    package: _Package,
    rels: dict[str, _Rel],
) -> dict[tuple[str, str], ET.Element]:
    """Collect placeholder geometry from the slide layout and master.

    Slide placeholders usually omit ``a:xfrm``; the layout often omits it as
    well and inherits from the master.  Resolve layout first, then fill any
    remaining keys from the master.
    """

    candidates: list[str] = []
    layout_rel = next(
        (item for item in rels.values() if item.type.endswith("/slideLayout")),
        None,
    )
    if layout_rel is not None and not layout_rel.external:
        candidates.append(layout_rel.target)
    master_rel = next(
        (item for item in rels.values() if item.type.endswith("/slideMaster")),
        None,
    )
    if master_rel is not None and not master_rel.external:
        candidates.append(master_rel.target)
    if layout_rel is not None and not layout_rel.external:
        master_rel = next(
            (
                item
                for item in package.rels(layout_rel.target).values()
                if item.type.endswith("/slideMaster")
            ),
            None,
        )
        if master_rel is not None and not master_rel.external:
            candidates.append(master_rel.target)

    result: dict[tuple[str, str], ET.Element] = {}
    for candidate in candidates:
        if candidate not in package.names:
            continue
        try:
            root = package.xml(candidate)
        except PptxNativeError:
            continue
        for shape in root.iter(_q("p", "sp")):
            key = _placeholder_key(shape)
            if key is None:
                continue
            xfrm = shape.find(f"./{_q('p', 'spPr')}/{_q('a', 'xfrm')}")
            if xfrm is None:
                continue
            result.setdefault(key, xfrm)
            result.setdefault((key[0], ""), xfrm)
    return result


def _shape_xfrm(
    shape: ET.Element,
    placeholder_xfrms: dict[tuple[str, str], ET.Element],
) -> ET.Element | None:
    xfrm = shape.find(f"./{_q('p', 'spPr')}/{_q('a', 'xfrm')}")
    if xfrm is None:
        xfrm = shape.find(_q("p", "xfrm"))
    if xfrm is not None:
        return xfrm
    key = _placeholder_key(shape)
    if key is not None:
        xfrm = placeholder_xfrms.get(key)
        if xfrm is None:
            xfrm = placeholder_xfrms.get((key[0], ""))
        return xfrm
    return None


def _append_shape(
    shape: ET.Element,
    package: _Package,
    theme: dict[str, str],
    fonts: dict[str, str],
    scale: float,
    stage_height: float,
    transform: _TRANSFORM,
    rels: dict[str, _Rel],
    placeholder_xfrms: dict[tuple[str, str], ET.Element],
    out: list[str],
) -> None:
    tag = shape.tag
    if tag == _q("p", "grpSp"):
        xfrm = shape.find(f"./{_q('p', 'grpSpPr')}/{_q('a', 'xfrm')}")
        left, top, width, height = _xfrm_box(xfrm, transform)
        if width <= 0 or height <= 0:
            return
        ch_off = xfrm.find(_q("a", "chOff")) if xfrm is not None else None
        ch_ext = xfrm.find(_q("a", "chExt")) if xfrm is not None else None
        child_x = _int(ch_off.get("x")) if ch_off is not None else 0
        child_y = _int(ch_off.get("y")) if ch_off is not None else 0
        child_w = _int(ch_ext.get("cx")) if ch_ext is not None else width
        child_h = _int(ch_ext.get("cy")) if ch_ext is not None else height
        child_scale_x = transform[2] * (width / child_w) if child_w else transform[2]
        child_scale_y = (
            transform[3] * (height / child_h) if child_h else transform[3]
        )
        child_transform = (
            left,
            top,
            child_scale_x,
            child_scale_y,
            float(child_x),
            float(child_y),
        )
        for child in shape:
            if child.tag.startswith(f"{{{NS['p']}}}"):
                _append_shape(
                    child,
                    package,
                    theme,
                    fonts,
                    scale,
                    stage_height,
                    child_transform,
                    rels,
                    placeholder_xfrms,
                    out,
                )
        return

    if tag not in {
        _q("p", "sp"),
        _q("p", "pic"),
        _q("p", "graphicFrame"),
        _q("p", "cxnSp"),
    }:
        return

    xfrm = _shape_xfrm(shape, placeholder_xfrms)
    left, top, width, height = _xfrm_box(xfrm, transform)
    if width <= 0 or height <= 0:
        return
    box = (left, top, width, height)
    transform_css = _rotation(xfrm)

    if tag in {_q("p", "cxnSp"),}:
        line_color = _fill_color(
            shape.find(f"./{_q('p', 'spPr')}/{_q('a', 'ln')}"),
            theme,
        )
        style = _shape_style(
            box,
            scale,
            stage_height,
            border=f"1.5px solid {line_color or '#94a3b8'}",
            transform=transform_css,
        )
        out.append(f'<div class="qop-shape" style="{style}"></div>')
        return

    if tag == _q("p", "pic"):
        blip = shape.find(f"./{_q('p', 'blipFill')}/{_q('a', 'blip')}")
        rel_id = blip.get(_q("r", "embed")) if blip is not None else None
        rel = rels.get(rel_id or "")
        uri = ""
        if rel and not rel.external and rel.target in package.names:
            try:
                uri, _, _ = image_data_uri(
                    package.read(rel.target),
                    package.content_type(rel.target),
                )
            except Exception as exc:
                logger.debug("Native PPTX image failed: %s", exc)
        if uri:
            name_node = shape.find(
                f"./{_q('p', 'nvPicPr')}/{_q('p', 'cNvPr')}",
            )
            alt = _escape(name_node.get("name") if name_node is not None else "")
            style = _shape_style(
                box,
                scale,
                stage_height,
                transform=transform_css,
            )
            out.append(
                f'<div class="qop-shape" style="{style}">'
                f'<img class="qop-shape-img" src="{uri}" alt="{alt}" '
                'style="width:100%;height:100%;object-fit:contain"></div>',
            )
        return

    if tag == _q("p", "graphicFrame"):
        graphic_data = shape.find(
            f"./{_q('a', 'graphic')}/{_q('a', 'graphicData')}",
        )
        if graphic_data is None:
            return
        uri = graphic_data.get("uri") or ""
        if uri.endswith("/table"):
            table = graphic_data.find(_q("a", "tbl"))
            if table is not None:
                style = _shape_style(
                    box,
                    scale,
                    stage_height,
                    transform=transform_css,
                )
                out.append(
                    f'<div class="qop-shape" style="{style}">'
                    f"{_table_html(table, _TextContext(theme, fonts, scale, rels), box)}"
                    "</div>",
                )
            return
        if uri.endswith("/chart"):
            chart_ref = graphic_data.find(_q("c", "chart"))
            rel_id = (
                chart_ref.get(_q("r", "id")) if chart_ref is not None else None
            )
            rel = rels.get(rel_id or "")
            style = _shape_style(
                box,
                scale,
                stage_height,
                transform=transform_css,
            )
            if rel is not None:
                out.append(
                    f'<div class="qop-shape" style="{style}">'
                    f"{_chart_html(package, rel, _TextContext(theme, fonts, scale, rels))}"
                    "</div>",
                )
            else:
                out.append(
                    f'<div class="qop-shape" style="{style}">'
                    '<div class="qop-empty">［图表］</div></div>',
                )
            return
        if "diagram" in uri:
            style = _shape_style(
                box,
                scale,
                stage_height,
                transform=transform_css,
            )
            out.append(
                f'<div class="qop-shape" style="{style}">'
                '<div class="qop-empty">［SmartArt］</div></div>',
            )
        return

    # Regular shape / placeholder.
    placeholder = _placeholder_of(shape)
    sp_pr = shape.find(_q("p", "spPr"))
    fill = _fill_color(sp_pr, theme)
    line = _fill_color(sp_pr.find(_q("a", "ln")), theme)
    border = f"1px solid {line}" if line else None
    tx_body = shape.find(_q("p", "txBody"))
    text_html = _text_frame_html(
        tx_body,
        _TextContext(
            theme=theme,
            fonts=fonts,
            scale=scale,
            rels=rels,
            placeholder=placeholder,
        ),
    )
    if not text_html and not fill and not border:
        return
    style = _shape_style(
        box,
        scale,
        stage_height,
        background=fill,
        border=border,
        transform=transform_css,
    )
    out.append(
        f'<div class="qop-shape" style="{style}">{text_html}</div>',
    )


# ---------------------------------------------------------------------------
# Slide helpers
# ---------------------------------------------------------------------------


def _slide_background(
    slide: ET.Element,
    theme: dict[str, str],
) -> str:
    bg_pr = slide.find(f"./{_q('p', 'cSld')}/{_q('p', 'bg')}/{_q('p', 'bgPr')}")
    color = _fill_color(bg_pr, theme)
    return color or ""


def _slide_texts(slide: ET.Element) -> list[str]:
    texts: list[str] = []
    for shape in slide.iter(_q("p", "sp")):
        if _placeholder_of(shape) == "sldNum":
            continue
        body = shape.find(_q("p", "txBody"))
        if body is None:
            continue
        for paragraph in body.findall(_q("a", "p")):
            text = "".join(
                node.text or "" for node in paragraph.iter(_q("a", "t"))
            ).strip()
            if text:
                texts.append(text)
    return texts


def _slide_title(slide: ET.Element) -> str:
    for shape in slide.findall(f"./{_q('p', 'cSld')}/{_q('p', 'spTree')}/{_q('p', 'sp')}"):
        placeholder = _placeholder_of(shape)
        if placeholder in {"title", "ctrTitle"}:
            texts = _slide_texts(shape)
            if texts:
                return texts[0]
    texts = _slide_texts(slide)
    return texts[0] if texts else ""


def _notes_text(package: _Package, slide_part: str) -> str:
    rels = package.rels(slide_part)
    notes_rel = next(
        (rel for rel in rels.values() if rel.type.endswith("/notesSlide")),
        None,
    )
    if notes_rel is None or notes_rel.target not in package.names:
        return ""
    try:
        notes = package.xml(notes_rel.target)
    except PptxNativeError:
        return ""
    lines: list[str] = []
    for shape in notes.iter(_q("p", "sp")):
        placeholder = _placeholder_of(shape)
        if placeholder in {"sldNum", "hdr", "ftr", "dt"}:
            continue
        body = shape.find(_q("p", "txBody"))
        if body is None:
            continue
        for paragraph in body.findall(_q("a", "p")):
            text = "".join(
                node.text or "" for node in paragraph.iter(_q("a", "t"))
            ).strip()
            if text:
                lines.append(text)
    return "\n".join(lines)


def _presentation_part(package: _Package) -> str:
    root_rels = package.rels("")
    for rel in root_rels.values():
        if rel.type.endswith("/officeDocument") and not rel.external:
            return rel.target
    if "ppt/presentation.xml" in package.names:
        return "ppt/presentation.xml"
    raise PptxNativeError("PPTX 中缺少 presentation.xml")


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------


def render_pptx_native(
    data: bytes,
    *,
    filename: str,
    theme: str = "auto",
    extra_warnings: list[str] | None = None,
) -> tuple[str, list[str]]:
    """Render a PPTX package using only the standard library."""

    package = _Package(data)
    try:
        presentation_part = _presentation_part(package)
        presentation = package.xml(presentation_part)
        size = presentation.find(_q("p", "sldSz"))
        slide_width = _int(size.get("cx")) if size is not None else 9144000
        slide_height = _int(size.get("cy")) if size is not None else 6858000
        if slide_width <= 0:
            slide_width = 9144000
        if slide_height <= 0:
            slide_height = 6858000
        scale = TARGET_SLIDE_WIDTH_PX / slide_width
        stage_height = round(slide_height * scale)

        colors, fonts = _load_theme(package)

        presentation_rels = package.rels(presentation_part)
        slide_parts: list[str] = []
        id_list = presentation.find(_q("p", "sldIdLst"))
        if id_list is not None:
            for sld_id in id_list.findall(_q("p", "sldId")):
                rel_id = sld_id.get(_q("r", "id"))
                rel = presentation_rels.get(rel_id or "")
                if rel is not None and not rel.external:
                    slide_parts.append(rel.target)
        if not slide_parts:
            slide_parts = sorted(
                (
                    name
                    for name in package.names
                    if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
                ),
                key=lambda name: int(re.search(r"(\d+)\.xml$", name).group(1)),  # type: ignore[union-attr]
            )
        slide_parts = slide_parts[:_MAX_SLIDES]

        slides_html: list[str] = []
        notes_html: list[str] = []
        outline: list[str] = []
        warnings: list[str] = list(extra_warnings or [])

        for index, part in enumerate(slide_parts, start=1):
            try:
                slide = package.xml(part)
            except PptxNativeError as exc:
                warnings.append(f"第 {index} 页解析失败，已跳过。")
                logger.info("Native PPTX slide failed: %s", exc)
                continue
            title = _slide_title(slide) or f"幻灯片 {index}"
            outline.append(
                f'<li><a href="#slide-{index}">{_escape(title)}</a></li>',
            )
            background = _slide_background(slide, colors)
            bg_style = f"background:{background};" if background else ""
            rels = package.rels(part)
            layout_xfrms = _placeholder_xfrms(package, rels)
            shapes: list[str] = []
            sp_tree = slide.find(f"./{_q('p', 'cSld')}/{_q('p', 'spTree')}")
            if sp_tree is not None:
                for child in sp_tree:
                    try:
                        _append_shape(
                            child,
                            package,
                            colors,
                            fonts,
                            scale,
                            stage_height,
                            (0.0, 0.0, 1.0, 1.0, 0.0, 0.0),
                            rels,
                            layout_xfrms,
                            shapes,
                        )
                    except Exception as exc:  # one bad shape must not kill deck
                        logger.debug("Native PPTX shape failed: %s", exc)
            slides_html.append(
                f'<section class="qop-slide" id="slide-{index}" '
                f'style="aspect-ratio:{slide_width}/{slide_height};">'
                f'<div class="qop-slide-inner" style="{bg_style}">'
                f'{"".join(shapes)}</div>'
                f'<span class="qop-slide-no">{index}</span>'
                "</section>",
            )
            notes = _notes_text(package, part)
            if notes:
                notes_html.append(
                    f'<details class="qop-notes" id="notes-{index}">'
                    f"<summary>第 {index} 页备注</summary>"
                    f"<p>{_escape(notes).replace(chr(10), '<br>')}</p></details>",
                )

        if not slides_html:
            warnings.append("演示文稿中没有可显示的幻灯片。")

        outline_html = (
            f'<nav class="qop-outline"><h2>幻灯片导航（共 {len(slides_html)} 页）</h2>'
            f"<ol>{''.join(outline)}</ol></nav>"
            if slides_html
            else ""
        )
        body = (
            f'{outline_html}<div class="qop-slides">{"".join(slides_html)}</div>'
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
    finally:
        package.close()
