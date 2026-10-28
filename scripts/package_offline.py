#!/usr/bin/env python3
"""Build a self-contained, offline-installable QwenPaw plugin zip.

The bundle vendors only pure-Python wheels (``requirements-offline.txt``) at
the plugin root. QwenPaw's per-plugin import namespace makes those packages
available to lazy imports inside ``converter/`` without touching ``sys.path``,
so the host never runs pip and no network access is required at install time.

Usage:
    python scripts/package_offline.py --out-dir release
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PLUGIN_ID = "qwenpaw-office-preview"
OFFLINE_REQUIREMENTS = ROOT / "requirements-offline.txt"
OFFLINE_TEMPLATE = ROOT / "packaging" / "OFFLINE-INSTALL.md"

# Files/directories that make up the installable plugin. Development-only
# content (frontend sources, tests, packaging scripts) is deliberately absent.
INCLUDE = [
    "plugin.py",
    "plugin.json",
    "converter",
    "dist",
    "docs",
    "README.md",
    "LICENSE",
]
EXCLUDE_NAMES = {
    "__pycache__",
    ".pytest_cache",
    ".ruff_cache",
    ".mypy_cache",
    ".DS_Store",
    "node_modules",
    ".venv",
}
BINARY_SUFFIXES = {".so", ".pyd", ".dylib", ".dll", ".a", ".lib"}


class PackagingError(RuntimeError):
    pass


def _source_date_epoch() -> int:
    raw = os.environ.get("SOURCE_DATE_EPOCH", "").strip()
    if raw.isdigit():
        return int(raw)
    try:
        return int(
            subprocess.check_output(
                ["git", "log", "-1", "--format=%ct"],
                cwd=ROOT,
                text=True,
            ).strip(),
        )
    except Exception:
        return int(time.time())


def _copy_tree(source: Path, destination: Path) -> None:
    if source.is_dir():
        for item in source.rglob("*"):
            if item.is_dir() or any(part in EXCLUDE_NAMES for part in item.parts):
                continue
            target = destination / item.relative_to(source)
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(item, target)
    else:
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)


def _install_vendor(staging: Path, python: str) -> None:
    if not OFFLINE_REQUIREMENTS.is_file():
        raise PackagingError(f"missing {OFFLINE_REQUIREMENTS.name}")
    command: list[str]
    uv = shutil.which("uv")
    if uv:
        command = [
            uv,
            "pip",
            "install",
            "--python",
            python,
            "--target",
            str(staging),
            "--no-compile",
            "--only-binary=:all:",
            "-r",
            str(OFFLINE_REQUIREMENTS),
        ]
    else:
        command = [
            python,
            "-m",
            "pip",
            "install",
            "--target",
            str(staging),
            "--no-compile",
            "--only-binary=:all:",
            "-r",
            str(OFFLINE_REQUIREMENTS),
        ]
    print("$", " ".join(command))
    completed = subprocess.run(command, check=False)
    if completed.returncode != 0:
        raise PackagingError("failed to install offline dependencies")


def _clean_vendor(staging: Path) -> tuple[list[Path], list[str]]:
    binaries: list[Path] = []
    for path in sorted(staging.rglob("*")):
        if path.is_dir():
            continue
        if path.suffix.lower() in BINARY_SUFFIXES:
            binaries.append(path.relative_to(staging))
    if binaries:
        joined = ", ".join(str(item) for item in binaries[:8])
        raise PackagingError(
            "offline bundle is not pure-Python; compiled artifacts found: "
            f"{joined}",
        )

    for path in list(staging.rglob("__pycache__")):
        shutil.rmtree(path, ignore_errors=True)
    for path in list(staging.rglob("*.pyc")):
        path.unlink(missing_ok=True)
    for name in ("bin", "Scripts", "node_modules"):
        shutil.rmtree(staging / name, ignore_errors=True)
    for name in (".lock", ".uv", ".gitignore"):
        (staging / name).unlink(missing_ok=True)

    packages: list[str] = []
    for entry in sorted(staging.iterdir()):
        if entry.is_dir() and (entry / "__init__.py").exists():
            packages.append(entry.name)
    return binaries, packages


def _write_manifest(staging: Path, version: str) -> dict:
    manifest_path = staging / "plugin.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if version:
        manifest["version"] = version
    manifest["dependencies"] = []
    meta = manifest.setdefault("meta", {})
    meta["offline_bundle"] = True
    meta["bundled_dependencies"] = ["mammoth", "openpyxl"]
    manifest_path.write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    return manifest


def _write_install_docs(staging: Path, version: str) -> None:
    text = OFFLINE_TEMPLATE.read_text(encoding="utf-8").replace(
        "{{VERSION}}",
        version,
    )
    (staging / "OFFLINE-INSTALL.md").write_text(text, encoding="utf-8")


def _write_build_info(
    staging: Path,
    manifest: dict,
    build_epoch: int,
    python: str,
) -> None:
    info = {
        "plugin_id": manifest.get("id", PLUGIN_ID),
        "version": manifest.get("version", ""),
        "offline": True,
        "build_time_utc": dt.datetime.fromtimestamp(
            build_epoch,
            tz=dt.timezone.utc,
        ).isoformat(),
        "python": python,
        "commit": os.environ.get("GITHUB_SHA", ""),
        "dependencies": {
            "mammoth": "1.13.0",
            "openpyxl": "3.1.5",
        },
        "pptx_renderer": "converter/pptx_native.py (pure Python)",
    }
    (staging / "BUILD-INFO.json").write_text(
        json.dumps(info, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def _zip_tree(root: Path, destination: Path, build_epoch: int) -> None:
    """Create a deterministic zip whose entries start with the root folder."""

    epoch = max(build_epoch, 315532800)  # zip timestamps start at 1980-01-01
    timestamp = time.gmtime(epoch)[:6]
    files = sorted(path for path in root.rglob("*") if path.is_file())
    with zipfile.ZipFile(
        destination,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=9,
    ) as archive:
        for path in files:
            relative = path.relative_to(root.parent).as_posix()
            info = zipfile.ZipInfo(relative, date_time=timestamp)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            info.create_system = 3
            archive.writestr(info, path.read_bytes())


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def build(out_dir: Path, version_override: str, python: str) -> dict:
    source_manifest = json.loads(
        (ROOT / "plugin.json").read_text(encoding="utf-8"),
    )
    version = version_override or source_manifest.get("version") or "0.0.0"
    if not (ROOT / "dist" / "index.js").is_file():
        raise PackagingError(
            "dist/index.js is missing; run `cd frontend && npm run build`",
        )

    out_dir.mkdir(parents=True, exist_ok=True)
    staging_root = out_dir / ".staging"
    shutil.rmtree(staging_root, ignore_errors=True)
    staging = staging_root / PLUGIN_ID
    staging.mkdir(parents=True)

    for item in INCLUDE:
        source = ROOT / item
        if not source.exists():
            raise PackagingError(f"missing plugin file: {item}")
        _copy_tree(source, staging / item)

    _install_vendor(staging, python)
    _, packages = _clean_vendor(staging)
    manifest = _write_manifest(staging, version)
    _write_install_docs(staging, version)
    build_epoch = _source_date_epoch()
    _write_build_info(staging, manifest, build_epoch, python)

    if (staging / "requirements.txt").exists():
        raise PackagingError("requirements.txt must not be shipped offline")

    zip_name = f"{PLUGIN_ID}-offline-{version}.zip"
    zip_path = out_dir / zip_name
    zip_path.unlink(missing_ok=True)
    _zip_tree(staging, zip_path, build_epoch)
    digest = _sha256(zip_path)
    checksum_path = out_dir / f"{zip_name}.sha256"
    checksum_path.write_text(
        f"{digest}  {zip_name}\n",
        encoding="utf-8",
    )

    summary = {
        "zip": str(zip_path),
        "sha256": digest,
        "size": zip_path.stat().st_size,
        "version": version,
        "packages": packages,
        "staging": str(staging),
        "files": sum(1 for item in staging.rglob("*") if item.is_file()),
    }
    print(json.dumps(summary, indent=2, ensure_ascii=False))
    return summary


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=ROOT / "dist-offline",
        help="Directory for the zip, checksum and staging tree.",
    )
    parser.add_argument(
        "--version",
        default="",
        help="Override the plugin version (defaults to plugin.json).",
    )
    parser.add_argument(
        "--python",
        default=sys.executable,
        help="Python interpreter used for uv/pip wheel resolution.",
    )
    args = parser.parse_args()
    try:
        build(args.out_dir.resolve(), args.version.strip(), args.python)
    except PackagingError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
