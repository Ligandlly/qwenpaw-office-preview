# -*- coding: utf-8 -*-
"""Small allowlist HTML sanitizer for third-party (mammoth) output.

The converter generates most of its HTML itself.  The one fragment that comes
from a library is the DOCX body produced by ``mammoth``.  Mammoth is
conservative, but a Word document can still carry hyperlinks such as
``javascript:...``, therefore everything that reaches an ``iframe`` goes
through this allowlist first.

Rendering happens inside ``<iframe sandbox="">`` (no scripts, no navigation),
so this is defence in depth rather than the only barrier.
"""

from __future__ import annotations

import html
import re
from html.parser import HTMLParser
from urllib.parse import urlsplit

_ALLOWED_TAGS = {
    "a",
    "b",
    "blockquote",
    "br",
    "caption",
    "code",
    "dd",
    "div",
    "dl",
    "dt",
    "em",
    "figcaption",
    "figure",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "hr",
    "i",
    "img",
    "li",
    "mark",
    "ol",
    "p",
    "pre",
    "s",
    "section",
    "small",
    "span",
    "strong",
    "sub",
    "sup",
    "table",
    "tbody",
    "td",
    "tfoot",
    "th",
    "thead",
    "tr",
    "u",
    "ul",
}

_VOID_TAGS = {"br", "hr", "img"}

# Tags whose *content* must disappear with them (not just the tag).
_DROP_CONTENT_TAGS = {
    "applet",
    "audio",
    "canvas",
    "embed",
    "iframe",
    "noscript",
    "object",
    "script",
    "style",
    "svg",
    "template",
    "video",
}

_GLOBAL_ATTRS = {"class", "style", "title"}

_TAG_ATTRS = {
    "a": {"href"},
    "img": {"src", "alt", "width", "height"},
    "td": {"colspan", "rowspan"},
    "th": {"colspan", "rowspan"},
}

_ALLOWED_STYLE_PROPS = {
    "background-color",
    "border-collapse",
    "color",
    "font-style",
    "font-weight",
    "height",
    "list-style-type",
    "margin-left",
    "max-width",
    "padding-left",
    "text-align",
    "text-decoration",
    "vertical-align",
    "width",
}

# ``expression()``, ``url()`` and CSS escapes have no business in a document
# preview; keep the property values to plain, boring declarations.
_UNSAFE_STYLE = re.compile(
    r"(?:url\s*\(|expression\s*\(|javascript:|vbscript:|@import|behavior\s*:|[\\<>])",
    re.IGNORECASE,
)
_SAFE_NUMBER = re.compile(r"^-?\d+(?:\.\d+)?(?:px|pt|em|rem|%|ex|ch)?$")
_SAFE_COLOR = re.compile(
    r"^(?:#[0-9a-fA-F]{3,8}|(?:rgb|rgba|hsl|hsla)\([0-9.,%\s]+\)|"
    r"[a-zA-Z]{3,20})$",
)
_SAFE_STYLE_KEYWORD = re.compile(r"^[a-zA-Z-]{3,32}$")

_ALLOWED_LINK_SCHEMES = {"http", "https", "mailto"}
_ALLOWED_IMAGE_SCHEMES = {"data"}


def _safe_url(value: str, schemes: set[str]) -> str | None:
    value = value.strip()
    if not value:
        return None
    if value.startswith("#"):
        # Same-document anchors are safe and useful inside the preview.
        return value
    try:
        parts = urlsplit(value)
    except ValueError:
        return None
    scheme = parts.scheme.lower()
    if scheme not in schemes:
        return None
    if scheme == "data" and not value.lower().startswith("data:image/"):
        return None
    return value


def _safe_style(value: str) -> str:
    declarations: list[str] = []
    for part in value.split(";"):
        if ":" not in part:
            continue
        prop, _, raw = part.partition(":")
        prop = prop.strip().lower()
        raw = raw.strip()
        if prop not in _ALLOWED_STYLE_PROPS or not raw:
            continue
        if _UNSAFE_STYLE.search(raw):
            continue
        if prop in {"width", "height", "max-width", "margin-left", "padding-left"}:
            if not _SAFE_NUMBER.match(raw):
                continue
        elif prop in {"color", "background-color"}:
            if not _SAFE_COLOR.match(raw):
                continue
        elif not _SAFE_STYLE_KEYWORD.match(raw):
            continue
        declarations.append(f"{prop}:{raw}")
    return ";".join(declarations)


def _clean_attrs(tag: str, attrs: list[tuple[str, str | None]]) -> str:
    allowed = _GLOBAL_ATTRS | _TAG_ATTRS.get(tag, set())
    rendered: list[str] = []
    for name, value in attrs:
        name = (name or "").lower()
        value = value if value is not None else ""
        if name not in allowed:
            continue
        if name == "href":
            safe = _safe_url(value, _ALLOWED_LINK_SCHEMES)
            if safe is None:
                continue
            value = safe
        elif name == "src":
            safe = _safe_url(value, _ALLOWED_IMAGE_SCHEMES)
            if safe is None:
                continue
            value = safe
        elif name in {"colspan", "rowspan", "width", "height"}:
            if not re.fullmatch(r"\d{1,4}", value.strip()):
                continue
            value = value.strip()
        elif name == "style":
            value = _safe_style(value)
            if not value:
                continue
        rendered.append(f' {name}="{html.escape(value, quote=True)}"')
    return "".join(rendered)


class _Sanitizer(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._out: list[str] = []
        self._open: list[str] = []
        self._drop_depth = 0

    def handle_starttag(self, tag, attrs) -> None:
        tag = tag.lower()
        if self._drop_depth:
            if tag in _DROP_CONTENT_TAGS:
                self._drop_depth += 1
            return
        if tag in _DROP_CONTENT_TAGS:
            self._drop_depth = 1
            return
        if tag not in _ALLOWED_TAGS:
            return
        self._out.append(f"<{tag}{_clean_attrs(tag, attrs)}>")
        if tag not in _VOID_TAGS:
            self._open.append(tag)

    def handle_startendtag(self, tag, attrs) -> None:
        tag = tag.lower()
        if self._drop_depth or tag not in _ALLOWED_TAGS:
            return
        self._out.append(f"<{tag}{_clean_attrs(tag, attrs)}>")
        if tag in _VOID_TAGS:
            return
        # An unclosed non-void element is closed immediately; the Markdown
        # style consumer (a browser) tolerates empty elements.
        self._out.append(f"</{tag}>")

    def handle_endtag(self, tag) -> None:
        tag = tag.lower()
        if self._drop_depth:
            if tag in _DROP_CONTENT_TAGS:
                self._drop_depth -= 1
            return
        if tag not in _ALLOWED_TAGS or tag in _VOID_TAGS:
            return
        if tag in self._open:
            # Close any elements the source left dangling.
            while self._open:
                pending = self._open.pop()
                self._out.append(f"</{pending}>")
                if pending == tag:
                    break

    def handle_data(self, data: str) -> None:
        if self._drop_depth:
            return
        self._out.append(html.escape(data, quote=False))

    def handle_comment(self, data: str) -> None:  # noqa: ARG002
        return

    def handle_decl(self, decl: str) -> None:  # noqa: ARG002
        return

    def handle_pi(self, data: str) -> None:  # noqa: ARG002
        return

    def result(self) -> str:
        while self._open:
            self._out.append(f"</{self._open.pop()}>")
        return "".join(self._out)


def sanitize_fragment(fragment: str) -> str:
    """Return *fragment* stripped down to the allowlist above."""

    parser = _Sanitizer()
    try:
        parser.feed(fragment)
        parser.close()
    except Exception:  # pragma: no cover - parser is lenient by design
        return html.escape(fragment, quote=False)
    return parser.result()
