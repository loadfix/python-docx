# pyright: reportPrivateUsage=false

"""Semantic diff between two |Document| instances (issue #75).

Compares the *content* of two documents — paragraph adds / removes /
modifications, table mutations, image counts, style additions — rather
than the raw XML, which would over-report whitespace and ordering noise
that makes no visible difference.

Three granularity levels:

* ``"structural"`` — paragraph add/remove/move only.
* ``"content"`` (default) — adds per-paragraph text changes on top of
  structural.
* ``"formatting"`` — adds style / font / colour changes on top of
  content.

Output formats:

* Python object — the default; iterate :attr:`SemanticDiff.changes`,
  read :attr:`SemanticDiff.summary`.
* :meth:`SemanticDiff.to_markdown` — PR-comment-friendly Markdown
  with a counts table and a per-change list.
* :meth:`SemanticDiff.to_html` — minimal HTML5 fragment, suitable for
  embedding in a web review UI.
* :meth:`SemanticDiff.to_word_track_changes` — emits a third
  |Document| whose paragraphs carry the diff as visible markers
  (``[INS]`` / ``[DEL]`` / ``[~MOD]``). Best-effort: full ``w:ins`` /
  ``w:del`` track-changes authoring is out of scope here — the marker
  format is human-readable and survives round-trip through Word.

Comparison strategy:

* Text-level — uses :class:`difflib.SequenceMatcher` over normalised
  paragraph-text strings, so whitespace and case noise are suppressed
  but real edits surface as ``replace`` / ``insert`` / ``delete``
  opcodes.
* Table-level — compares cell-text grids; reports row/col adds /
  removes / cell modifications inside the table.
* Image-level — compares :attr:`Document.inline_shapes` length and
  the SHA-1 of each inline picture's binary part. Only counts /
  presence are reported; pixel-diffing is out of scope.
* Style-level — compares :attr:`Document.styles` membership by style
  ID; new / removed / renamed styles are reported.

The module pulls in only :mod:`difflib` from stdlib; no NLP or
heavy-deps creep allowed.

.. versionadded:: 2026.05.13
"""

from __future__ import annotations

import difflib
import hashlib
import html
import re
from dataclasses import dataclass, field
from typing import (
    TYPE_CHECKING,
    Any,
    Iterable,
    List,
    Optional,
    Sequence,
    Tuple,
    Union,
)

if TYPE_CHECKING:
    from docx.document import Document
    from docx.styles.style import BaseStyle
    from docx.table import Table, _Cell
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run


__all__ = [
    "Change",
    "SemanticDiff",
    "compute_diff",
    "VALID_LEVELS",
]


VALID_LEVELS: Tuple[str, ...] = ("structural", "content", "formatting")

# -- Whitespace-collapse pattern used when normalising paragraph text
# -- for the structural / content comparison. Difflib's similarity
# -- ratio is sensitive to whitespace, so collapsing here keeps the
# -- comparison focussed on meaningful edits. Matching diffs in Word
# -- typically present whitespace-only re-flow as a non-change. --
_WS_RE = re.compile(r"\s+")


def _normalise(text: str) -> str:
    """Collapse runs of whitespace to a single space and strip."""
    return _WS_RE.sub(" ", text or "").strip()


# ---------------------------------------------------------------------------
# Change record
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Change:
    """One semantic diff finding.

    Attributes:
        kind: ``"paragraph_added"`` | ``"paragraph_removed"`` |
            ``"paragraph_modified"`` | ``"paragraph_moved"`` |
            ``"table_modified"`` | ``"table_added"`` |
            ``"table_removed"`` | ``"image_added"`` |
            ``"image_removed"`` | ``"style_added"`` |
            ``"style_removed"`` | ``"style_changed"`` |
            ``"formatting_changed"``.
        target: Locator string — typically ``"paragraph[i]"`` or
            ``"table[i]"`` or a style name.
        before: Pre-edit value (string, dict, or ``None`` for adds).
        after: Post-edit value (string, dict, or ``None`` for removes).
        detail: Free-form message; populated for table cell and
            formatting findings where the locator alone is ambiguous.
    """

    kind: str
    target: str
    before: Any = None
    after: Any = None
    detail: Optional[str] = None

    def to_dict(self) -> dict:
        """Return a JSON-serialisable view of this change."""
        return {
            "kind": self.kind,
            "target": self.target,
            "before": self.before,
            "after": self.after,
            "detail": self.detail,
        }


# ---------------------------------------------------------------------------
# SemanticDiff result
# ---------------------------------------------------------------------------


@dataclass
class SemanticDiff:
    """Aggregate result of comparing two |Document| instances.

    Returned by :meth:`docx.document.Document.diff`.
    """

    level: str
    changes: List[Change] = field(default_factory=list)

    # ---- counts -------------------------------------------------------

    @property
    def summary(self) -> dict:
        """Return a counts dictionary keyed by change-kind family.

        Example::

            {'paragraphs_added': 3, 'paragraphs_removed': 1,
             'paragraphs_modified': 7, 'tables_modified': 1,
             'images_added': 0, 'styles_changed': 0,
             'total_changes': 12}
        """
        counts = {
            "paragraphs_added": 0,
            "paragraphs_removed": 0,
            "paragraphs_modified": 0,
            "paragraphs_moved": 0,
            "tables_added": 0,
            "tables_removed": 0,
            "tables_modified": 0,
            "images_added": 0,
            "images_removed": 0,
            "styles_added": 0,
            "styles_removed": 0,
            "styles_changed": 0,
            "formatting_changed": 0,
        }
        for c in self.changes:
            key = {
                "paragraph_added": "paragraphs_added",
                "paragraph_removed": "paragraphs_removed",
                "paragraph_modified": "paragraphs_modified",
                "paragraph_moved": "paragraphs_moved",
                "table_added": "tables_added",
                "table_removed": "tables_removed",
                "table_modified": "tables_modified",
                "image_added": "images_added",
                "image_removed": "images_removed",
                "style_added": "styles_added",
                "style_removed": "styles_removed",
                "style_changed": "styles_changed",
                "formatting_changed": "formatting_changed",
            }.get(c.kind)
            if key is not None:
                counts[key] += 1
        counts["total_changes"] = len(self.changes)
        return counts

    def __len__(self) -> int:
        return len(self.changes)

    def __iter__(self):
        return iter(self.changes)

    def __bool__(self) -> bool:
        return bool(self.changes)

    # ---- helpers ------------------------------------------------------

    def filter(self, *kinds: str) -> List[Change]:
        """Return changes whose ``kind`` matches any of `kinds`."""
        wanted = set(kinds)
        return [c for c in self.changes if c.kind in wanted]

    # ---- exporters ----------------------------------------------------

    def to_markdown(self, max_per_kind: int = 25) -> str:
        """Return a PR-comment-friendly Markdown summary.

        ``max_per_kind`` caps the number of detail lines printed per
        change-kind so the comment stays scannable on huge diffs; the
        counts table is always complete.
        """
        s = self.summary
        lines: List[str] = []
        lines.append("### Document diff (`level=%s`)" % self.level)
        lines.append("")
        lines.append("| Kind | Count |")
        lines.append("| --- | ---: |")
        for label, key in (
            ("Paragraphs added", "paragraphs_added"),
            ("Paragraphs removed", "paragraphs_removed"),
            ("Paragraphs modified", "paragraphs_modified"),
            ("Paragraphs moved", "paragraphs_moved"),
            ("Tables added", "tables_added"),
            ("Tables removed", "tables_removed"),
            ("Tables modified", "tables_modified"),
            ("Images added", "images_added"),
            ("Images removed", "images_removed"),
            ("Styles added", "styles_added"),
            ("Styles removed", "styles_removed"),
            ("Styles changed", "styles_changed"),
            ("Formatting changed", "formatting_changed"),
        ):
            if s.get(key, 0):
                lines.append("| %s | %d |" % (label, s[key]))
        lines.append("| **Total changes** | **%d** |" % s["total_changes"])
        lines.append("")

        # -- per-change detail, grouped by kind --
        if self.changes:
            grouped: dict = {}
            for c in self.changes:
                grouped.setdefault(c.kind, []).append(c)
            for kind in sorted(grouped):
                bucket = grouped[kind]
                lines.append("#### %s (%d)" % (kind, len(bucket)))
                lines.append("")
                for c in bucket[:max_per_kind]:
                    lines.append("- " + _md_change_line(c))
                if len(bucket) > max_per_kind:
                    lines.append(
                        "- _... %d more elided ..._" % (len(bucket) - max_per_kind)
                    )
                lines.append("")
        return "\n".join(lines).rstrip() + "\n"

    def to_html(self) -> str:
        """Return a minimal self-contained HTML5 fragment.

        Suitable for embedding in a web review UI. Text content is
        HTML-escaped at every leaf to guard against XSS from document
        content.
        """
        s = self.summary
        out: List[str] = []
        out.append('<div class="docx-semantic-diff">')
        out.append(
            '<h3>Document diff <small>(level=<code>%s</code>)</small></h3>'
            % html.escape(self.level)
        )
        out.append("<table><thead><tr><th>Kind</th><th>Count</th></tr></thead><tbody>")
        for label, key in (
            ("Paragraphs added", "paragraphs_added"),
            ("Paragraphs removed", "paragraphs_removed"),
            ("Paragraphs modified", "paragraphs_modified"),
            ("Paragraphs moved", "paragraphs_moved"),
            ("Tables added", "tables_added"),
            ("Tables removed", "tables_removed"),
            ("Tables modified", "tables_modified"),
            ("Images added", "images_added"),
            ("Images removed", "images_removed"),
            ("Styles added", "styles_added"),
            ("Styles removed", "styles_removed"),
            ("Styles changed", "styles_changed"),
            ("Formatting changed", "formatting_changed"),
        ):
            if s.get(key, 0):
                out.append(
                    "<tr><td>%s</td><td>%d</td></tr>"
                    % (html.escape(label), s[key])
                )
        out.append(
            "<tr><td><strong>Total</strong></td><td><strong>%d</strong></td></tr>"
            % s["total_changes"]
        )
        out.append("</tbody></table>")

        if self.changes:
            out.append("<ul>")
            for c in self.changes:
                css = _kind_css_class(c.kind)
                out.append(
                    '<li class="%s"><strong>%s</strong> %s%s</li>'
                    % (
                        css,
                        html.escape(c.kind),
                        html.escape(c.target),
                        _html_change_payload(c),
                    )
                )
            out.append("</ul>")
        out.append("</div>")
        return "\n".join(out)

    def to_word_track_changes(self):
        """Emit a third |Document| with the diff rendered as visible markers.

        Best-effort: full ``w:ins`` / ``w:del`` track-changes authoring
        is out of scope for this exporter. Each change is rendered as
        a paragraph carrying a human-readable marker prefix
        (``[INS] ...``, ``[DEL] ...``, ``[~MOD] before -> after``,
        etc.) so the resulting document can be opened in Word and
        skimmed.

        Returns a fresh :class:`docx.document.Document` constructed from
        the package default template. The caller is responsible for
        saving it.

        Limitations
        -----------
        * Markers are text-level; a Word reviewer cannot accept /
          reject them as native revisions.
        * Original formatting (run colour, font) is *not* preserved;
          the output is a structural summary, not a styled rendering.
        * Tables / images / styles findings collapse to one paragraph
          per finding — granular cell-level review is available only
          via :attr:`changes`.

        .. versionadded:: 2026.05.13
        """
        from docx import Document as _open

        doc = _open()
        doc.add_heading("Document diff (level=%s)" % self.level, level=1)
        s = self.summary
        doc.add_paragraph(
            "Total changes: %d (paragraphs_added=%d, paragraphs_removed=%d, "
            "paragraphs_modified=%d, tables_modified=%d, images_added=%d, "
            "styles_changed=%d)"
            % (
                s["total_changes"],
                s["paragraphs_added"],
                s["paragraphs_removed"],
                s["paragraphs_modified"],
                s["tables_modified"],
                s["images_added"],
                s["styles_changed"],
            )
        )
        for c in self.changes:
            marker = {
                "paragraph_added": "[INS]",
                "paragraph_removed": "[DEL]",
                "paragraph_modified": "[~MOD]",
                "paragraph_moved": "[MOV]",
                "table_added": "[TBL+]",
                "table_removed": "[TBL-]",
                "table_modified": "[TBL~]",
                "image_added": "[IMG+]",
                "image_removed": "[IMG-]",
                "style_added": "[STY+]",
                "style_removed": "[STY-]",
                "style_changed": "[STY~]",
                "formatting_changed": "[FMT~]",
            }.get(c.kind, "[?]")
            body = _plain_change_line(c)
            doc.add_paragraph("%s %s %s" % (marker, c.target, body))
        return doc


# ---------------------------------------------------------------------------
# Markdown / HTML helpers
# ---------------------------------------------------------------------------


def _md_change_line(c: Change) -> str:
    """Format one Change as a Markdown bullet body."""
    if c.kind in ("paragraph_added", "image_added", "style_added", "table_added"):
        return "`%s` -> %s" % (c.target, _md_value(c.after))
    if c.kind in (
        "paragraph_removed",
        "image_removed",
        "style_removed",
        "table_removed",
    ):
        return "`%s` (was %s)" % (c.target, _md_value(c.before))
    if c.kind in (
        "paragraph_modified",
        "style_changed",
        "formatting_changed",
        "table_modified",
    ):
        if c.detail:
            return "`%s`: %s" % (c.target, c.detail)
        return "`%s`: %s -> %s" % (
            c.target,
            _md_value(c.before),
            _md_value(c.after),
        )
    if c.kind == "paragraph_moved":
        return "`%s` (was %s)" % (c.target, _md_value(c.before))
    return "`%s` %s -> %s" % (c.target, _md_value(c.before), _md_value(c.after))


def _md_value(v: Any) -> str:
    if v is None:
        return "_(none)_"
    if isinstance(v, str):
        s = v.replace("|", "\\|")
        if len(s) > 80:
            s = s[:77] + "..."
        return "`" + s + "`"
    return "`" + repr(v) + "`"


def _html_change_payload(c: Change) -> str:
    if c.detail:
        return ": " + html.escape(c.detail)
    if c.before is None and c.after is not None:
        return " -> " + html.escape(_short(c.after))
    if c.after is None and c.before is not None:
        return " (was " + html.escape(_short(c.before)) + ")"
    if c.before is not None or c.after is not None:
        return (
            " "
            + html.escape(_short(c.before))
            + " -> "
            + html.escape(_short(c.after))
        )
    return ""


def _kind_css_class(kind: str) -> str:
    if kind.endswith("_added"):
        return "diff-added"
    if kind.endswith("_removed"):
        return "diff-removed"
    if kind.endswith("_modified") or kind.endswith("_changed"):
        return "diff-modified"
    if kind.endswith("_moved"):
        return "diff-moved"
    return "diff-other"


def _short(v: Any, limit: int = 80) -> str:
    s = "" if v is None else (v if isinstance(v, str) else repr(v))
    if len(s) > limit:
        s = s[: limit - 3] + "..."
    return s


def _plain_change_line(c: Change) -> str:
    if c.detail:
        return c.detail
    if c.before is None and c.after is not None:
        return _short(c.after, 200)
    if c.after is None and c.before is not None:
        return _short(c.before, 200)
    if c.before is not None or c.after is not None:
        return _short(c.before, 100) + " -> " + _short(c.after, 100)
    return ""


# ---------------------------------------------------------------------------
# Snapshots — extract a comparable view of each document
# ---------------------------------------------------------------------------


def _paragraph_snapshot(p: "Paragraph") -> dict:
    """Return a hashable, comparable view of a Paragraph."""
    style_name: Optional[str] = None
    try:
        style = p.style
        if style is not None:
            style_name = getattr(style, "name", None)
    except Exception:  # pragma: no cover - defensive
        style_name = None
    runs_fmt: List[Tuple] = []
    for r in p.runs:
        runs_fmt.append(_run_format_signature(r))
    return {
        "text": p.text or "",
        "norm": _normalise(p.text or ""),
        "style": style_name,
        "runs_fmt": tuple(runs_fmt),
    }


def _run_format_signature(r: "Run") -> Tuple:
    """A small signature of a Run's direct formatting."""
    font = r.font
    color_hex: Optional[str] = None
    try:
        rgb = font.color.rgb if font.color is not None else None
        color_hex = str(rgb) if rgb is not None else None
    except Exception:  # pragma: no cover - defensive
        color_hex = None
    size = None
    try:
        size = int(font.size) if font.size is not None else None
    except Exception:  # pragma: no cover - defensive
        size = None
    return (
        getattr(font, "name", None),
        size,
        font.bold,
        font.italic,
        getattr(font, "underline", None),
        color_hex,
    )


def _table_snapshot(t: "Table") -> dict:
    """Return a comparable view of a Table — cell-text grid + dimensions."""
    grid: List[List[str]] = []
    try:
        rows = list(t.rows)
    except Exception:  # pragma: no cover - defensive
        rows = []
    for row in rows:
        row_cells: List[str] = []
        try:
            for cell in row.cells:
                row_cells.append(_normalise(cell.text or ""))
        except Exception:  # pragma: no cover - defensive
            pass
        grid.append(row_cells)
    style_name: Optional[str] = None
    try:
        if t.style is not None:
            style_name = getattr(t.style, "name", None)
    except Exception:  # pragma: no cover - defensive
        pass
    return {
        "rows": len(grid),
        "cols": max((len(r) for r in grid), default=0),
        "grid": grid,
        "style": style_name,
    }


def _image_signatures(doc: "Document") -> List[str]:
    """Return a list of SHA-1 digests of every inline picture's bytes.

    Falls back to a stable per-shape locator when the binary part is
    unreachable (corrupt relationship) so identity-by-hash never throws.
    """
    sigs: List[str] = []
    try:
        shapes = list(doc.inline_shapes)
    except Exception:  # pragma: no cover - defensive
        return sigs
    for idx, shape in enumerate(shapes):
        digest = "shape-%d" % idx
        try:
            blip = shape._inline.graphic.graphicData.pic.blipFill.blip
            rId = blip.embed
            part = doc.part.related_parts.get(rId) if rId else None
            if part is not None and getattr(part, "blob", None) is not None:
                digest = hashlib.sha1(part.blob).hexdigest()
        except Exception:  # pragma: no cover - defensive
            pass
        sigs.append(digest)
    return sigs


def _style_snapshot(doc: "Document") -> dict:
    """Return a {style_name: signature} mapping for the document's styles.

    The signature is a tuple of the style's font name / size / bold /
    italic / colour, so a font change on the same style shows up as
    ``style_changed`` rather than only style add/remove.
    """
    out: dict = {}
    try:
        styles = doc.styles
    except Exception:  # pragma: no cover - defensive
        return out
    for style in styles:
        name = getattr(style, "name", None)
        if name is None:
            continue
        out[name] = _style_signature(style)
    return out


def _style_signature(style: "BaseStyle") -> Tuple:
    """Hashable signature of a Style's user-visible font properties."""
    font = getattr(style, "font", None)
    if font is None:
        return ()
    name = getattr(font, "name", None)
    size = None
    try:
        size = int(font.size) if font.size is not None else None
    except Exception:  # pragma: no cover - defensive
        size = None
    bold = getattr(font, "bold", None)
    italic = getattr(font, "italic", None)
    underline = getattr(font, "underline", None)
    color_hex: Optional[str] = None
    try:
        rgb = font.color.rgb if font.color is not None else None
        color_hex = str(rgb) if rgb is not None else None
    except Exception:  # pragma: no cover - defensive
        pass
    return (name, size, bold, italic, underline, color_hex)


# ---------------------------------------------------------------------------
# Top-level driver
# ---------------------------------------------------------------------------


def compute_diff(
    old: "Document", new: "Document", level: str = "content"
) -> SemanticDiff:
    """Compute a :class:`SemanticDiff` between two documents.

    Args:
        old: The pre-edit |Document|.
        new: The post-edit |Document|.
        level: One of :data:`VALID_LEVELS` —
            ``"structural"`` / ``"content"`` (default) /
            ``"formatting"``.

    Raises:
        ValueError: If `level` is not one of the recognised tokens.
    """
    if level not in VALID_LEVELS:
        raise ValueError(
            "level must be one of %r, got %r" % (VALID_LEVELS, level)
        )

    diff = SemanticDiff(level=level)

    # -- 1. Paragraphs -------------------------------------------------
    _diff_paragraphs(old, new, level, diff)

    # -- 2. Tables -----------------------------------------------------
    _diff_tables(old, new, level, diff)

    # -- 3. Images -----------------------------------------------------
    _diff_images(old, new, diff)

    # -- 4. Styles (formatting level only — at content / structural a
    # -- style-only edit isn't reported) ------------------------------
    if level == "formatting":
        _diff_styles(old, new, diff)

    return diff


def _diff_paragraphs(
    old: "Document",
    new: "Document",
    level: str,
    diff: SemanticDiff,
) -> None:
    old_paragraphs = list(old.paragraphs)
    new_paragraphs = list(new.paragraphs)
    old_snaps = [_paragraph_snapshot(p) for p in old_paragraphs]
    new_snaps = [_paragraph_snapshot(p) for p in new_paragraphs]

    old_keys = [s["norm"] for s in old_snaps]
    new_keys = [s["norm"] for s in new_snaps]

    matcher = difflib.SequenceMatcher(a=old_keys, b=new_keys, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            if level == "formatting":
                # -- look for style/run-formatting drift on equal-text pairs --
                for offset in range(i2 - i1):
                    o = old_snaps[i1 + offset]
                    n = new_snaps[j1 + offset]
                    if o["style"] != n["style"]:
                        diff.changes.append(
                            Change(
                                kind="formatting_changed",
                                target="paragraph[%d]" % (j1 + offset),
                                before={"style": o["style"]},
                                after={"style": n["style"]},
                                detail="style %r -> %r"
                                % (o["style"], n["style"]),
                            )
                        )
                    if o["runs_fmt"] != n["runs_fmt"]:
                        diff.changes.append(
                            Change(
                                kind="formatting_changed",
                                target="paragraph[%d]" % (j1 + offset),
                                before={"runs_fmt": list(o["runs_fmt"])},
                                after={"runs_fmt": list(n["runs_fmt"])},
                                detail="run formatting changed",
                            )
                        )
            continue
        if tag == "delete":
            for k in range(i1, i2):
                diff.changes.append(
                    Change(
                        kind="paragraph_removed",
                        target="paragraph[%d]" % k,
                        before=old_snaps[k]["text"],
                        after=None,
                    )
                )
        elif tag == "insert":
            for k in range(j1, j2):
                diff.changes.append(
                    Change(
                        kind="paragraph_added",
                        target="paragraph[%d]" % k,
                        before=None,
                        after=new_snaps[k]["text"],
                    )
                )
        elif tag == "replace":
            # -- pair up at the level=structural we'd emit add/remove
            # -- separately, but at content/formatting we surface
            # -- pair-aligned modifications which read better. --
            old_block = list(range(i1, i2))
            new_block = list(range(j1, j2))
            common = min(len(old_block), len(new_block))
            if level == "structural":
                # -- structural mode: pure add / remove, no per-text edits
                for k in old_block:
                    diff.changes.append(
                        Change(
                            kind="paragraph_removed",
                            target="paragraph[%d]" % k,
                            before=old_snaps[k]["text"],
                        )
                    )
                for k in new_block:
                    diff.changes.append(
                        Change(
                            kind="paragraph_added",
                            target="paragraph[%d]" % k,
                            after=new_snaps[k]["text"],
                        )
                    )
                continue
            for offset in range(common):
                o_idx = old_block[offset]
                n_idx = new_block[offset]
                diff.changes.append(
                    Change(
                        kind="paragraph_modified",
                        target="paragraph[%d]" % n_idx,
                        before=old_snaps[o_idx]["text"],
                        after=new_snaps[n_idx]["text"],
                    )
                )
            for k in old_block[common:]:
                diff.changes.append(
                    Change(
                        kind="paragraph_removed",
                        target="paragraph[%d]" % k,
                        before=old_snaps[k]["text"],
                    )
                )
            for k in new_block[common:]:
                diff.changes.append(
                    Change(
                        kind="paragraph_added",
                        target="paragraph[%d]" % k,
                        after=new_snaps[k]["text"],
                    )
                )


def _diff_tables(
    old: "Document",
    new: "Document",
    level: str,
    diff: SemanticDiff,
) -> None:
    try:
        old_tables = list(old.tables)
        new_tables = list(new.tables)
    except Exception:  # pragma: no cover - defensive
        return
    old_snaps = [_table_snapshot(t) for t in old_tables]
    new_snaps = [_table_snapshot(t) for t in new_tables]

    # -- positional alignment by index — same as paragraphs but with
    # -- coarser comparators since SequenceMatcher on grids is
    # -- noisy. --
    common = min(len(old_snaps), len(new_snaps))
    for i in range(common):
        o = old_snaps[i]
        n = new_snaps[i]
        if o == n:
            continue
        details: List[str] = []
        if o["rows"] != n["rows"]:
            details.append("rows %d -> %d" % (o["rows"], n["rows"]))
        if o["cols"] != n["cols"]:
            details.append("cols %d -> %d" % (o["cols"], n["cols"]))
        if level == "formatting" and o["style"] != n["style"]:
            details.append(
                "style %r -> %r" % (o["style"], n["style"])
            )
        cell_changes = _table_cell_changes(o["grid"], n["grid"])
        if cell_changes:
            details.append("%d cell(s) changed" % cell_changes)
        if not details:
            # -- only an unobserved structural difference; record a
            # -- generic finding so the change isn't silently dropped --
            details.append("table content changed")
        diff.changes.append(
            Change(
                kind="table_modified",
                target="table[%d]" % i,
                before={
                    "rows": o["rows"],
                    "cols": o["cols"],
                    "style": o["style"],
                },
                after={
                    "rows": n["rows"],
                    "cols": n["cols"],
                    "style": n["style"],
                },
                detail="; ".join(details),
            )
        )
    for i in range(common, len(old_snaps)):
        diff.changes.append(
            Change(
                kind="table_removed",
                target="table[%d]" % i,
                before={"rows": old_snaps[i]["rows"], "cols": old_snaps[i]["cols"]},
            )
        )
    for i in range(common, len(new_snaps)):
        diff.changes.append(
            Change(
                kind="table_added",
                target="table[%d]" % i,
                after={"rows": new_snaps[i]["rows"], "cols": new_snaps[i]["cols"]},
            )
        )


def _table_cell_changes(
    old_grid: Sequence[Sequence[str]], new_grid: Sequence[Sequence[str]]
) -> int:
    """Return the count of cell positions whose normalised text differs."""
    rows = max(len(old_grid), len(new_grid))
    diffs = 0
    for r in range(rows):
        old_row = old_grid[r] if r < len(old_grid) else []
        new_row = new_grid[r] if r < len(new_grid) else []
        cols = max(len(old_row), len(new_row))
        for c in range(cols):
            o = old_row[c] if c < len(old_row) else ""
            n = new_row[c] if c < len(new_row) else ""
            if o != n:
                diffs += 1
    return diffs


def _diff_images(old: "Document", new: "Document", diff: SemanticDiff) -> None:
    old_sigs = _image_signatures(old)
    new_sigs = _image_signatures(new)
    old_set = list(old_sigs)
    for sig in new_sigs:
        if sig in old_set:
            old_set.remove(sig)
        else:
            diff.changes.append(
                Change(kind="image_added", target="image[+]", after=sig)
            )
    for sig in old_set:
        diff.changes.append(
            Change(kind="image_removed", target="image[-]", before=sig)
        )


def _diff_styles(old: "Document", new: "Document", diff: SemanticDiff) -> None:
    old_styles = _style_snapshot(old)
    new_styles = _style_snapshot(new)
    old_names = set(old_styles)
    new_names = set(new_styles)
    for name in sorted(new_names - old_names):
        diff.changes.append(
            Change(
                kind="style_added",
                target="style[%s]" % name,
                after=name,
            )
        )
    for name in sorted(old_names - new_names):
        diff.changes.append(
            Change(
                kind="style_removed",
                target="style[%s]" % name,
                before=name,
            )
        )
    for name in sorted(old_names & new_names):
        if old_styles[name] != new_styles[name]:
            diff.changes.append(
                Change(
                    kind="style_changed",
                    target="style[%s]" % name,
                    before=old_styles[name],
                    after=new_styles[name],
                    detail="font signature changed",
                )
            )
