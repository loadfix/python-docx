#!/usr/bin/env python3
"""Prepare a new release of python-docx (fork).

Bumps ``__version__`` in ``src/docx/__init__.py``, prompts for a new
``HISTORY.rst`` entry, validates the HISTORY formatting, then runs
``pytest`` and ``pyright`` so the release can't be tagged with a
broken tree.

Usage
-----

    python scripts/prepare_release.py <new-version>

The tool is deliberately interactive — it is intended to be run by a
human cutting a release, not by CI. It makes edits in place but does
not tag or push; that is done manually after review with::

    git diff                         # sanity-check the bump
    git add -p                       # stage
    git commit -m 'chore(release): bump to X.Y.Z'
    git tag -a vX.Y.Z -m 'X.Y.Z'
    git push origin master --follow-tags

The ``.github/workflows/release.yml`` tag trigger then takes over.

Version policy
--------------

The fork uses CalVer (``YYYY.MM.patch``) — e.g. ``2026.05.7``. The
script validates the new version string against that shape but does
not enforce that it is strictly greater than the current one (so a
hotfix to an older month is allowed).
"""

from __future__ import annotations

import argparse
import datetime as _dt
import re
import subprocess
import sys
from pathlib import Path

# --- Configuration (library-specific) --------------------------------

PROJECT_ROOT = Path(__file__).resolve().parent.parent
VERSION_FILE = PROJECT_ROOT / "src" / "docx" / "__init__.py"
HISTORY_FILE = PROJECT_ROOT / "HISTORY.rst"
LIBRARY_NAME = "python-docx"

# Commands run to validate the tree before allowing the release to
# proceed. Both must exit 0. If a user wants to skip a step (e.g. no
# pyright available locally) pass ``--skip-checks``.
CHECK_COMMANDS: list[list[str]] = [
    ["pytest", "tests/", "-q"],
    ["pyright", "src/"],
]

# --- Shared logic (keep in sync across sibling repos) ----------------

CALVER_RE = re.compile(r"^(?P<year>20\d{2})\.(?P<month>0[1-9]|1[0-2])\.(?P<patch>\d+)$")
VERSION_LINE_RE = re.compile(
    r'^__version__\s*=\s*["\'](?P<v>[^"\']+)["\']', re.MULTILINE
)


class ReleaseError(RuntimeError):
    """Raised when the release preparation fails in a user-visible way."""


def read_current_version() -> str:
    text = VERSION_FILE.read_text()
    m = VERSION_LINE_RE.search(text)
    if not m:
        raise ReleaseError(
            f"Could not find __version__ = '...' in {VERSION_FILE}"
        )
    return m.group("v")


def write_new_version(new_version: str) -> None:
    text = VERSION_FILE.read_text()
    new_text, n = VERSION_LINE_RE.subn(
        lambda _: f'__version__ = "{new_version}"', text, count=1
    )
    if n != 1:
        raise ReleaseError(
            f"Failed to rewrite __version__ in {VERSION_FILE}"
        )
    VERSION_FILE.write_text(new_text)


def validate_calver(v: str) -> None:
    m = CALVER_RE.match(v)
    if not m:
        raise ReleaseError(
            f"Version {v!r} does not match CalVer shape YYYY.MM.patch "
            "(e.g. 2026.05.7)."
        )


def prompt_history_entry(new_version: str) -> str:
    print()
    print(f"Enter a HISTORY.rst entry for {new_version}.")
    print("Paste the bulleted body (without the version header / underline).")
    print("Terminate with a line containing only a single '.' (period).")
    print()
    lines: list[str] = []
    while True:
        try:
            line = input()
        except EOFError:
            break
        if line.strip() == ".":
            break
        lines.append(line)
    body = "\n".join(lines).strip()
    if not body:
        raise ReleaseError("HISTORY entry body must not be empty.")
    return body


def render_history_entry(new_version: str, body: str, title: str | None = None) -> str:
    """Render a release block matching the existing HISTORY.rst style."""
    header = new_version if not title else f"{new_version} — {title}"
    underline = "+" * len(header)
    today = _dt.date.today().isoformat()
    return f"\n{header}\n{underline}\n\nReleased: {today}\n\n{body}\n\n"


def insert_history_entry(block: str) -> None:
    """Insert ``block`` directly below the "Release History" heading."""
    text = HISTORY_FILE.read_text()
    # The HISTORY.rst files across the three libraries all lead with::
    #     .. :changelog:
    #
    #     Release History
    #     ---------------
    #
    # We insert directly after the underline.
    marker = re.compile(
        r"(^Release History\s*\n-+\s*\n)", re.MULTILINE
    )
    m = marker.search(text)
    if not m:
        raise ReleaseError(
            f"Could not find 'Release History' heading in {HISTORY_FILE}"
        )
    insert_at = m.end()
    new_text = text[:insert_at] + block + text[insert_at:]
    HISTORY_FILE.write_text(new_text)


def validate_history_format(new_version: str) -> None:
    """Re-parse HISTORY.rst to confirm the new entry is well-formed."""
    text = HISTORY_FILE.read_text()
    # Look for the new version header followed by an underline of the
    # same length and a "Released: YYYY-MM-DD" line within the next
    # 10 lines.
    header_re = re.compile(
        rf"^{re.escape(new_version)}(\s+—\s+.*)?$\n(?P<under>\+{{3,}})\s*$",
        re.MULTILINE,
    )
    m = header_re.search(text)
    if not m:
        raise ReleaseError(
            f"HISTORY.rst does not contain a well-formed header for "
            f"{new_version} (expected 'VERSION' then a line of '+' "
            "characters at least as long)."
        )
    header_len = len(new_version)
    if m.group("under") and len(m.group("under")) < header_len:
        raise ReleaseError(
            f"HISTORY.rst underline for {new_version} is shorter than the "
            "header text — RST will render it as a paragraph."
        )
    # Check "Released: YYYY-MM-DD" line shows up within the following 10 lines.
    tail = text[m.end() :].splitlines()[:10]
    if not any(
        re.match(r"Released:\s+\d{4}-\d{2}-\d{2}\s*$", ln) for ln in tail
    ):
        raise ReleaseError(
            f"HISTORY.rst entry for {new_version} is missing a "
            "'Released: YYYY-MM-DD' line within the first 10 lines "
            "of the block."
        )


def run_check(cmd: list[str]) -> None:
    print(f"  $ {' '.join(cmd)}")
    try:
        subprocess.run(cmd, check=True, cwd=PROJECT_ROOT)
    except FileNotFoundError as exc:
        raise ReleaseError(
            f"Required tool not found on PATH: {exc.filename!r}. "
            "Install the dev extras ('pip install -e .[dev]') or pass "
            "--skip-checks if you really know what you're doing."
        ) from exc
    except subprocess.CalledProcessError as exc:
        raise ReleaseError(
            f"Check failed (exit {exc.returncode}): {' '.join(cmd)}"
        ) from exc


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=f"Prepare a {LIBRARY_NAME} release.")
    p.add_argument(
        "new_version",
        help="New CalVer version string (e.g. 2026.05.7). "
        "Must match YYYY.MM.patch.",
    )
    p.add_argument(
        "--title",
        default=None,
        help="Optional short title appended to the version header "
        '(e.g. --title "Foo support"). Renders as "VERSION — Foo support".',
    )
    p.add_argument(
        "--skip-checks",
        action="store_true",
        help="Skip pytest + pyright. Don't use this unless you have "
        "already run them by hand.",
    )
    p.add_argument(
        "--history-body",
        default=None,
        help="HISTORY.rst entry body. If omitted, the tool prompts "
        "for it interactively (terminate with a line containing '.').",
    )
    args = p.parse_args(argv)

    try:
        validate_calver(args.new_version)
        current = read_current_version()
        if current == args.new_version:
            raise ReleaseError(
                f"New version {args.new_version} is identical to current."
            )
        print(f"Bumping {LIBRARY_NAME}: {current} -> {args.new_version}")

        body = args.history_body or prompt_history_entry(args.new_version)
        block = render_history_entry(args.new_version, body, title=args.title)

        print("Writing version bump...")
        write_new_version(args.new_version)
        print("Inserting HISTORY.rst entry...")
        insert_history_entry(block)
        print("Validating HISTORY.rst formatting...")
        validate_history_format(args.new_version)

        if args.skip_checks:
            print("Skipping pytest + pyright (per --skip-checks).")
        else:
            print("Running pre-release checks...")
            for cmd in CHECK_COMMANDS:
                run_check(cmd)

        print()
        print("Release prepared. Next:")
        print("  git diff")
        print("  git add -p")
        print(f"  git commit -m 'chore(release): bump to {args.new_version}'")
        print(f"  git tag -a v{args.new_version} -m '{args.new_version}'")
        print("  git push origin master --follow-tags")
        return 0
    except ReleaseError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
