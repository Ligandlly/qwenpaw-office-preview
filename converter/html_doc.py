# -*- coding: utf-8 -*-
"""Self-contained HTML documents used by the preview iframe.

Every renderer returns one complete HTML document with inline CSS and inline
(data-URI) images so the Console can put it in a sandboxed ``iframe`` without
loading any external resource.
"""

from __future__ import annotations

import html

_LIGHT_VARS = """
  --qop-fg: #1f2328;
  --qop-fg-muted: #5b6472;
  --qop-bg: #eef1f5;
  --qop-paper: #ffffff;
  --qop-border: #d7dce3;
  --qop-accent: #2563eb;
  --qop-soft: #f4f6f9;
  --qop-code-bg: #f0f2f5;
"""

_DARK_VARS = """
  --qop-fg: #e6e9ef;
  --qop-fg-muted: #a0a8b6;
  --qop-bg: #14171c;
  --qop-paper: #1c2027;
  --qop-border: #313743;
  --qop-accent: #7aa2ff;
  --qop-soft: #232830;
  --qop-code-bg: #262c35;
"""

_BASE_CSS = r"""
:root {
  color-scheme: light dark;
""" + _LIGHT_VARS + r"""
}
@media (prefers-color-scheme: dark) {
  :root {
""" + _DARK_VARS + r"""
  }
}
/* The Console can force its own theme, overriding the OS preference. */
html.qop-theme-light {
  color-scheme: light;
""" + _LIGHT_VARS + r"""
}
html.qop-theme-dark {
  color-scheme: dark;
""" + _DARK_VARS + r"""
}
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
html {
  background: var(--qop-bg);
  scrollbar-width: none;
}
html::-webkit-scrollbar { width: 0; height: 0; }
* {
  scrollbar-color: var(--qop-border) transparent;
  scrollbar-width: thin;
}
*::-webkit-scrollbar { width: 8px; height: 8px; }
*::-webkit-scrollbar-thumb {
  background: var(--qop-border);
  border-radius: 4px;
}
*::-webkit-scrollbar-track { background: transparent; }
body {
  min-height: 100vh;
  padding: 22px 16px 48px;
  background: var(--qop-bg);
  color: var(--qop-fg);
  font-family: "Segoe UI", "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei",
    "Noto Sans CJK SC", system-ui, -apple-system, sans-serif;
  font-size: 15px;
  line-height: 1.7;
  -webkit-font-smoothing: antialiased;
  overflow-wrap: anywhere;
}
.qop-doc {
  max-width: 940px;
  margin: 0 auto;
  padding: 44px 52px 56px;
  background: var(--qop-paper);
  border: 1px solid var(--qop-border);
  border-radius: 10px;
  box-shadow: 0 10px 30px rgba(15, 23, 42, 0.08), 0 2px 6px rgba(15, 23, 42, 0.04);
}
.qop-doc > :first-child { margin-top: 0; }
.qop-doc > :last-child { margin-bottom: 0; }
h1, h2, h3, h4, h5, h6 {
  margin: 1.5em 0 0.6em;
  line-height: 1.3;
  font-weight: 650;
  color: var(--qop-fg);
}
h1 { font-size: 1.9em; }
h2 { font-size: 1.55em; }
h3 { font-size: 1.28em; }
h4 { font-size: 1.1em; }
h5, h6 { font-size: 1em; }
p { margin: 0.72em 0; }
a { color: var(--qop-accent); text-decoration: none; }
a:hover { text-decoration: underline; }
ul, ol { margin: 0.6em 0; padding-left: 2em; }
li { margin: 0.22em 0; }
blockquote {
  margin: 1em 0;
  padding: 0.4em 1em;
  border-left: 3px solid var(--qop-border);
  color: var(--qop-fg-muted);
  background: var(--qop-soft);
}
code, pre {
  font-family: "SFMono-Regular", Consolas, "Liberation Mono", Menlo, monospace;
  font-size: 0.9em;
  background: var(--qop-code-bg);
  border-radius: 5px;
}
code { padding: 0.1em 0.35em; }
pre { padding: 0.9em 1.1em; overflow: auto; }
pre code { padding: 0; background: none; }
img { max-width: 100%; height: auto; }
hr { border: 0; border-top: 1px solid var(--qop-border); margin: 2em 0; }
table {
  border-collapse: collapse;
  width: 100%;
  margin: 1em 0;
  font-size: 0.94em;
}
th, td {
  border: 1px solid var(--qop-border);
  padding: 0.42em 0.62em;
  vertical-align: top;
  text-align: left;
}
th { background: var(--qop-soft); font-weight: 600; }
.qop-subtitle { color: var(--qop-fg-muted); font-size: 1.05em; margin-top: -0.4em; }
.qop-meta {
  max-width: 940px;
  margin: 0 auto 14px;
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  color: var(--qop-fg-muted);
  font-size: 12px;
}
.qop-chip {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  padding: 2px 9px;
  border: 1px solid var(--qop-border);
  border-radius: 999px;
  background: var(--qop-paper);
}
.qop-warning {
  max-width: 940px;
  margin: 0 auto 14px;
  padding: 9px 13px;
  border: 1px solid #f0c36d;
  border-radius: 8px;
  background: rgba(240, 195, 109, 0.12);
  color: var(--qop-fg);
  font-size: 13px;
}
.qop-empty { color: var(--qop-fg-muted); font-style: italic; }
@media (max-width: 640px) {
  body { padding: 10px 6px 28px; }
  .qop-doc { padding: 24px 18px 30px; border-radius: 6px; }
}
@media print {
  body { background: #fff; padding: 0; }
  .qop-doc { border: 0; box-shadow: none; max-width: none; padding: 0; }
  .qop-meta, .qop-warning { display: none; }
}
"""

_PPTX_CSS = r"""
body.qop-pptx {
  background: var(--qop-bg);
  padding-bottom: 40px;
}
.qop-slides {
  width: 100%;
  max-width: 1000px;
  margin: 0 auto;
  container-type: inline-size;
}
.qop-slide {
  position: relative;
  width: 100%;
  height: auto;
  margin: 0 auto 26px;
  background: var(--qop-paper);
  border: 1px solid var(--qop-border);
  border-radius: 8px;
  box-shadow: 0 8px 22px rgba(15, 23, 42, 0.1);
  overflow: hidden;
}
.qop-slide-inner { position: absolute; inset: 0; }
.qop-shape { position: absolute; overflow: visible; }
.qop-shape-text { white-space: pre-wrap; word-break: break-word; }
.qop-shape-text p { margin: 0; }
.qop-shape-text .qop-run { white-space: pre-wrap; }
.qop-shape-img { display: block; }
.qop-slide-no {
  position: absolute;
  right: 10px;
  bottom: 6px;
  font-size: clamp(9px, 1.15cqw, 12px);
  color: var(--qop-fg-muted);
  z-index: 2;
}
.qop-notes {
  max-width: 1000px;
  margin: -14px auto 26px;
  padding: 10px 14px;
  border: 1px dashed var(--qop-border);
  border-radius: 0 0 8px 8px;
  background: var(--qop-paper);
  color: var(--qop-fg-muted);
  font-size: 13px;
}
.qop-notes summary { cursor: pointer; color: var(--qop-fg); }
.qop-outline {
  max-width: 1000px;
  margin: 0 auto 18px;
  padding: 12px 16px;
  background: var(--qop-paper);
  border: 1px solid var(--qop-border);
  border-radius: 8px;
}
.qop-outline h2 { margin: 0 0 8px; font-size: 13px; text-transform: uppercase; letter-spacing: .06em; color: var(--qop-fg-muted); }
.qop-outline ol { margin: 0; padding-left: 1.6em; columns: 2; column-gap: 26px; }
.qop-outline li { margin: 0.18em 0; }
.qop-outline a { color: var(--qop-fg); }
.qop-chart-table { font-size: 11px; }
.qop-chart-table th, .qop-chart-table td { padding: 2px 6px; }
.qop-swatch { display: inline-block; width: 9px; height: 9px; border-radius: 2px; margin-right: 5px; vertical-align: -1px; }
@media (max-width: 1030px) {
  .qop-slides, .qop-outline, .qop-notes { max-width: none; }
  .qop-slide { width: 100%; }
}
"""

_XLSX_CSS = r"""
body.qop-xlsx { background: var(--qop-bg); }
.qop-workbook { max-width: 100%; margin: 0 auto; }
.qop-sheet {
  margin: 0 auto 26px;
  background: var(--qop-paper);
  border: 1px solid var(--qop-border);
  border-radius: 8px;
  box-shadow: 0 6px 18px rgba(15, 23, 42, 0.08);
  overflow: hidden;
}
.qop-sheet h2 {
  margin: 0;
  padding: 10px 14px;
  font-size: 14px;
  background: var(--qop-soft);
  border-bottom: 1px solid var(--qop-border);
}
.qop-sheet-scroll { overflow: auto; max-height: 78vh; }
table.qop-grid {
  border-collapse: separate;
  border-spacing: 0;
  margin: 0;
  font-size: 12.5px;
  font-variant-numeric: tabular-nums;
}
table.qop-grid th, table.qop-grid td {
  border-right: 1px solid var(--qop-border);
  border-bottom: 1px solid var(--qop-border);
  padding: 3px 7px;
  min-width: 56px;
  max-width: 420px;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}
table.qop-grid thead th {
  position: sticky;
  top: 0;
  z-index: 2;
  background: var(--qop-soft);
  font-weight: 600;
  color: var(--qop-fg-muted);
  text-align: center;
}
table.qop-grid tbody th {
  position: sticky;
  left: 0;
  z-index: 1;
  background: var(--qop-soft);
  font-weight: 500;
  color: var(--qop-fg-muted);
  text-align: right;
}
table.qop-grid thead th.qop-corner { left: 0; z-index: 3; }
table.qop-grid td.qop-num { text-align: right; }
table.qop-grid td.qop-bool { text-align: center; }
table.qop-grid td.qop-formula { color: var(--qop-fg-muted); font-family: Consolas, monospace; }
"""


def document(
    body: str,
    *,
    title: str,
    extra_css: str = "",
    body_class: str = "",
    meta: str = "",
    warnings: str = "",
    theme: str = "auto",
) -> str:
    """Wrap *body* in a complete, self-contained HTML document."""

    safe_title = html.escape(title or "Office preview", quote=True)
    meta_html = f'<div class="qop-meta">{meta}</div>' if meta else ""
    warning_html = (
        f'<div class="qop-warning">{warnings}</div>' if warnings else ""
    )
    class_attr = f' class="{body_class}"' if body_class else ""
    html_class = ""
    if theme == "dark":
        html_class = ' class="qop-theme-dark"'
    elif theme == "light":
        html_class = ' class="qop-theme-light"'
    return (
        "<!doctype html>\n"
        f'<html lang="zh-CN"{html_class}>\n<head>\n'
        '<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        '<meta name="color-scheme" content="light dark">\n'
        f"<title>{safe_title}</title>\n"
        f"<style>{_BASE_CSS}{extra_css}</style>\n"
        "</head>\n"
        f"<body{class_attr}>\n"
        f"{meta_html}{warning_html}"
        f"{body}\n"
        "</body>\n</html>"
    )


def pptx_document(
    body: str,
    *,
    title: str,
    meta: str = "",
    warnings: str = "",
    theme: str = "auto",
) -> str:
    return document(
        body,
        title=title,
        extra_css=_PPTX_CSS,
        body_class="qop-pptx",
        meta=meta,
        warnings=warnings,
        theme=theme,
    )


def xlsx_document(
    body: str,
    *,
    title: str,
    meta: str = "",
    warnings: str = "",
    theme: str = "auto",
) -> str:
    return document(
        body,
        title=title,
        extra_css=_XLSX_CSS,
        body_class="qop-xlsx",
        meta=meta,
        warnings=warnings,
        theme=theme,
    )
