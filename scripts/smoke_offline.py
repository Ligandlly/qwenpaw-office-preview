#!/usr/bin/env python3
"""Smoke-test an unpacked offline plugin tree.

The script imports the converter from the plugin root (where the bundled pure
Python dependencies live), blocks python-pptx / python-docx so the offline code
paths are exercised, and converts the committed fixtures.
"""

from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / "tests" / "fixtures"

EXPECTED = {
    "sample.docx": ["Offline Sample Report", "offline smoke tests", "<table"],
    "sample.xlsx": ["Data", "Notes", "35.0%", "¥12,345.67", "Merged Title"],
    "sample.pptx": [
        "Offline Sample Deck",
        "First part",
        "Table Slide",
        "Speaker note for offline test",
        "data:image/",
    ],
    "sample.csv": ["Alice", "90"],
}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("plugin_root", type=Path)
    args = parser.parse_args()
    plugin_root = args.plugin_root.resolve()

    manifest = plugin_root / "plugin.json"
    if not manifest.is_file():
        print(f"FAIL: {manifest} is missing", file=sys.stderr)
        return 1
    if (plugin_root / "requirements.txt").exists():
        print("FAIL: offline bundle must not contain requirements.txt", file=sys.stderr)
        return 1

    # Put the bundle first so the vendored packages win over anything already
    # installed in the test interpreter.
    sys.path.insert(0, str(plugin_root))

    for name in ("mammoth", "openpyxl"):
        spec = importlib.util.find_spec(name)
        if spec is None or not Path(spec.origin or "").is_relative_to(plugin_root):
            origin = spec.origin if spec else "not found"
            print(
                f"FAIL: {name} was not loaded from the bundle: {origin}",
                file=sys.stderr,
            )
            return 1

    # Force the offline code paths even on a developer machine that happens to
    # have python-pptx/python-docx installed.
    sys.modules["pptx"] = None  # type: ignore[assignment]
    sys.modules["docx"] = None  # type: ignore[assignment]
    sys.modules["lxml"] = None  # type: ignore[assignment]

    from converter.service import convert_office

    for filename, markers in EXPECTED.items():
        data = (FIXTURES / filename).read_bytes()
        result = convert_office(filename, data)
        html = result["html"]
        for marker in markers:
            if marker not in html:
                print(
                    f"FAIL: {filename} is missing marker {marker!r}",
                    file=sys.stderr,
                )
                return 1
        print(
            f"ok  {filename:14s} kind={result['kind']:5s} "
            f"html={len(html):6d} bytes",
        )

    import mammoth
    import openpyxl

    print(f"mammoth={mammoth.__file__}")
    print(f"openpyxl={openpyxl.__file__}")
    print("offline smoke test passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
