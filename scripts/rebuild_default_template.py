"""Rebuild ``src/docx/templates/default.docx`` from the unzipped template.

The unzipped template lives at ``src/docx/templates/default-docx-template/``
and is the source of truth for the default document blob that
``Document()`` loads when called without an argument. The zipped
``default.docx`` next to it is a derived artefact — it has repeatedly
drifted out of sync (most recently in 2026.05.2 when the Word-2024
namespace declarations were added to ``word/document.xml`` but the zip
was not regenerated).

Running this script repackages every file under the template directory
into ``default.docx`` deterministically — members are written in sorted
order with a fixed timestamp so the output is byte-reproducible.

Usage::

    python scripts/rebuild_default_template.py

The script has no side-effects beyond (re)writing
``src/docx/templates/default.docx``.
"""

from __future__ import annotations

import os
import sys
import zipfile
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parent.parent
SOURCE_DIR = REPO_ROOT / "src" / "docx" / "templates" / "default-docx-template"
OUTPUT_PATH = REPO_ROOT / "src" / "docx" / "templates" / "default.docx"

# Fixed timestamp — the zip format's minimum representable date. Matches
# the timestamp the reproducible-save writer emits.
_FIXED_DATE_TIME = (1980, 1, 1, 0, 0, 0)


def _collect_members(source_dir: Path) -> list[tuple[str, Path]]:
    """Return ``(archive_name, filesystem_path)`` pairs sorted by archive name."""
    members: list[tuple[str, Path]] = []
    for path in source_dir.rglob("*"):
        if not path.is_file():
            continue
        # Use forward slashes — zip archives are POSIX-ish regardless of host OS.
        rel = path.relative_to(source_dir).as_posix()
        members.append((rel, path))
    members.sort(key=lambda item: item[0])
    return members


def rebuild(source_dir: Path = SOURCE_DIR, output_path: Path = OUTPUT_PATH) -> Path:
    """Rebuild ``output_path`` from the tree rooted at ``source_dir``.

    Returns the absolute output path. Overwrites any existing zip at
    ``output_path``.
    """
    if not source_dir.is_dir():
        raise FileNotFoundError(f"template source directory missing: {source_dir}")

    members = _collect_members(source_dir)
    if not members:
        raise RuntimeError(f"no files found under {source_dir}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    # ``ZipFile`` in write mode truncates any existing file at the path.
    with zipfile.ZipFile(output_path, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for arcname, path in members:
            info = zipfile.ZipInfo(filename=arcname, date_time=_FIXED_DATE_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            # Normalise external_attr — regular file, 0o644 permissions.
            info.external_attr = (0o100644 & 0xFFFF) << 16
            with path.open("rb") as f:
                zf.writestr(info, f.read())

    return output_path


def main() -> int:
    output = rebuild()
    size = output.stat().st_size
    print(f"Wrote {output} ({size} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
