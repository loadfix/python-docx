"""Document outline / heading-tree helpers for python-docx.

Closes #71. Builds a hierarchical view of a document's heading
paragraphs so callers (LLM agents, navigation panes, summarisers,
slicers) can reason about structure without re-implementing the
~50 lines of style-walk + nest-by-level boilerplate every time.

The outline mirrors the role of pptx's ``deck.summarize()`` /
``skeleton()``: a compact, JSON-serialisable snapshot of the
document scaffold.

The exporter is deliberately layout-agnostic — python-docx has no
rendering engine, so page numbers cannot be computed and are
omitted from the output. Word's cached ``Pages`` value (read from
``docProps/app.xml`` via :class:`docx.statistics.DocumentStatistics`)
is surfaced as ``total_pages_estimated`` when available so the
caller can present *something* for the whole document.

Public surface:

* :class:`OutlineNode` — a single section (heading + nested children).
* :class:`Outline` — the top-level wrapper exposing ``walk()`` / ``to_dict()``.
* :func:`build_outline` — pure-function constructor used by
  :meth:`docx.document.Document.outline`.
* :func:`slice_document` — the implementation behind
  :meth:`docx.document.Document.slice`.

.. versionadded:: 2026.05.7
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Iterator, Sequence

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph


# -- "Heading 1" .. "Heading 9", case-insensitive. Title is treated as
# -- level 0 — Word's default template uses the ``Title`` style for the
# -- top-level cover heading and that's how :meth:`Document.add_heading`
# -- emits ``level=0``. --
_HEADING_RE = re.compile(r"^heading\s+([1-9])$", re.IGNORECASE)
_TITLE_NAMES = frozenset({"title"})


def _heading_level(paragraph: "Paragraph") -> "int | None":
    """Return the heading level of `paragraph`, or |None| if not a heading.

    ``Title`` is treated as level 0 (mirrors :meth:`Document.add_heading`
    semantics where ``level=0`` writes ``style="Title"``). ``Heading N``
    for ``N`` in 1..9 is treated as level ``N``. Anything else returns
    |None|.
    """
    style = paragraph.style
    if style is None:
        return None
    name = getattr(style, "name", None)
    if name is None:
        return None
    name = name.strip()
    if name.lower() in _TITLE_NAMES:
        return 0
    match = _HEADING_RE.match(name)
    if match is None:
        return None
    return int(match.group(1))


def _word_count(text: str) -> int:
    """Return the whitespace-token count of `text` (matches str.split)."""
    if not text:
        return 0
    return len(text.split())


def _stable_id(level: int, text: str, paragraph_index: int) -> str:
    """Return a deterministic short id for a heading.

    Word does not require a heading paragraph to be backed by a
    bookmark, so we synthesise an 8-char SHA-1 prefix derived from the
    heading's level, text, and document-order index. The result is
    stable across runs for the same content but does *not* round-trip
    through edits — adding paragraphs above the heading shifts its
    paragraph index and therefore its id. The caller treats the id as
    a snapshot key, not a persistent identifier.
    """
    payload = f"{level}\x1f{text}\x1f{paragraph_index}".encode("utf-8")
    return hashlib.sha1(payload).hexdigest()[:8]


@dataclass
class OutlineNode:
    """One heading-rooted section of a document outline.

    Attributes:

    - ``level`` — the heading's outline level (0 for ``Title``, 1..9
      for ``Heading 1``..``Heading 9``).
    - ``text`` — the heading paragraph's plain text.
    - ``paragraph_index`` — the heading paragraph's position in
      :attr:`docx.document.Document.paragraphs` (the body story only).
    - ``id`` — a stable 8-char hex id derived from ``level``, ``text``,
      and ``paragraph_index``. Treat as a snapshot key, not a
      persistent identifier (see :func:`_stable_id`).
    - ``word_count`` — number of whitespace-delimited tokens between
      this heading and the next heading at the same-or-shallower
      level (the section's body text). Includes the heading text
      itself.
    - ``children`` — list of nested |OutlineNode| objects whose
      ``level`` is strictly deeper than this one's.

    Page numbers are intentionally omitted: python-docx has no layout
    engine and cannot compute them. Callers needing approximate page
    positions can read the cached value from
    :attr:`Document.statistics.pages` (Word's last-saved count).

    .. versionadded:: 2026.05.7
    """

    level: int
    text: str
    paragraph_index: int
    id: str = ""
    word_count: int = 0
    children: "list[OutlineNode]" = field(default_factory=list)

    def walk(self) -> "Iterator[OutlineNode]":
        """Yield this node and every descendant in depth-first document order.

        .. versionadded:: 2026.05.7
        """
        yield self
        for child in self.children:
            yield from child.walk()

    def to_dict(self) -> "dict[str, object]":
        """Return a plain dict suitable for ``json.dumps``.

        Children are recursively converted. Mirrors the schema in
        the issue example so callers can pass the result straight to
        an LLM tool.

        .. versionadded:: 2026.05.7
        """
        return {
            "id": self.id,
            "heading": self.text,
            "level": self.level,
            "paragraph_index": self.paragraph_index,
            "word_count": self.word_count,
            "children": [c.to_dict() for c in self.children],
        }


@dataclass
class Outline:
    """Top-level outline of a |Document|.

    Wraps a list of root-level |OutlineNode| sections plus a few
    document-wide aggregates (``title``, ``total_paragraphs``,
    ``total_pages_estimated``). Constructed by :func:`build_outline`
    and surfaced by :meth:`docx.document.Document.outline`.

    .. versionadded:: 2026.05.7
    """

    sections: "list[OutlineNode]"
    title: "str | None" = None
    total_paragraphs: int = 0
    total_pages_estimated: "int | None" = None

    def walk(self) -> "Iterator[OutlineNode]":
        """Yield every |OutlineNode| in depth-first document order.

        .. versionadded:: 2026.05.7
        """
        for section in self.sections:
            yield from section.walk()

    def __iter__(self) -> "Iterator[OutlineNode]":
        return self.walk()

    def __len__(self) -> int:
        return sum(1 for _ in self.walk())

    def to_dict(self) -> "dict[str, object]":
        """Return a JSON-serialisable dict.

        Schema::

            {
                "title": str | None,
                "total_paragraphs": int,
                "total_pages_estimated": int | None,
                "sections": [OutlineNode.to_dict(), ...],
            }

        .. versionadded:: 2026.05.7
        """
        return {
            "title": self.title,
            "total_paragraphs": self.total_paragraphs,
            "total_pages_estimated": self.total_pages_estimated,
            "sections": [s.to_dict() for s in self.sections],
        }

    def find(self, heading: str) -> "OutlineNode | None":
        """Return the first |OutlineNode| whose ``text`` equals `heading`.

        Comparison is exact (case-sensitive, whitespace-stripped on
        both sides). Returns |None| when no heading matches.

        .. versionadded:: 2026.05.7
        """
        target = heading.strip()
        for node in self.walk():
            if node.text.strip() == target:
                return node
        return None


def build_outline(document: "Document") -> Outline:
    """Return an :class:`Outline` snapshot of `document`'s heading tree.

    Walks :attr:`Document.paragraphs` once. For each paragraph that
    resolves to a heading style (``Title`` → level 0, ``Heading N``
    → level ``N``), builds an :class:`OutlineNode` and nests it under
    the most-recent shallower heading. Paragraphs between two
    headings (the section body) contribute to the preceding
    heading's ``word_count``; the heading's own text is included in
    that count.

    The outline's ``title`` is sourced from the first ``Title``-styled
    paragraph if present, falling back to the document's
    core-properties ``title`` when one is set, then |None|.

    ``total_pages_estimated`` is populated from the cached
    ``docProps/app.xml`` ``<Pages>`` value (Word's last-saved page
    count) — see :attr:`Document.statistics`. Returns |None| when
    the value is unavailable.

    .. versionadded:: 2026.05.7
    """
    paragraphs = list(document.paragraphs)
    nodes_by_index: "list[tuple[int, OutlineNode]]" = []
    title: "str | None" = None

    # -- pass 1: build flat list of (paragraph_index, node) for each heading --
    for idx, paragraph in enumerate(paragraphs):
        level = _heading_level(paragraph)
        if level is None:
            continue
        text = paragraph.text or ""
        node = OutlineNode(
            level=level,
            text=text,
            paragraph_index=idx,
            id=_stable_id(level, text, idx),
        )
        nodes_by_index.append((idx, node))
        if level == 0 and title is None:
            title = text or None

    # -- pass 2: compute word_count per node (sum body paragraphs in section) --
    n = len(nodes_by_index)
    for i, (start_idx, node) in enumerate(nodes_by_index):
        end_idx = (
            nodes_by_index[i + 1][0] if i + 1 < n else len(paragraphs)
        )
        words = 0
        for p_idx in range(start_idx, end_idx):
            words += _word_count(paragraphs[p_idx].text or "")
        node.word_count = words

    # -- pass 3: nest using a stack keyed by level. Each stack entry is
    # -- (level, node). A new heading of level L pops entries >= L, then
    # -- attaches under the new top of stack. --
    roots: "list[OutlineNode]" = []
    stack: "list[OutlineNode]" = []
    for _, node in nodes_by_index:
        while stack and stack[-1].level >= node.level:
            stack.pop()
        if stack:
            stack[-1].children.append(node)
        else:
            roots.append(node)
        stack.append(node)

    # -- title fallback: core properties --
    if title is None:
        try:
            core_title = document.core_properties.title
        except Exception:
            core_title = None
        if isinstance(core_title, str) and core_title.strip():
            title = core_title

    # -- pages: read cached <Pages> from app.xml (no layout engine here) --
    pages_estimated: "int | None" = None
    try:
        stats = document.statistics
        raw_pages = getattr(stats, "pages", None)
        if isinstance(raw_pages, int) and not isinstance(raw_pages, bool):
            pages_estimated = raw_pages
    except Exception:
        pages_estimated = None

    return Outline(
        sections=roots,
        title=title,
        total_paragraphs=len(paragraphs),
        total_pages_estimated=pages_estimated,
    )


def slice_document(
    document: "Document",
    start: "str | OutlineNode",
    end: "str | OutlineNode | None" = None,
) -> "Document":
    """Return a new |Document| containing the paragraphs of one section.

    `start` selects the heading that opens the slice — either a
    heading's exact text (matched against :class:`OutlineNode.text`)
    or an :class:`OutlineNode` returned by :meth:`Outline.find` /
    :meth:`Outline.walk`. The slice runs from `start`'s paragraph
    (inclusive) up to but not including `end`'s paragraph; when
    `end` is |None| the slice runs to the end of the document.

    The new document is created from the same default template as
    :func:`docx.Document` (no argument) and paragraphs are copied
    using :meth:`Document.append_paragraph`, which rewires images,
    hyperlinks, and style references along the way.

    Raises :class:`ValueError` when `start` (or `end`, if a string)
    does not match any heading in `document`.

    .. versionadded:: 2026.05.7
    """
    from docx.api import Document as _DocumentFactory

    outline = build_outline(document)

    def _resolve(target: "str | OutlineNode") -> OutlineNode:
        if isinstance(target, OutlineNode):
            return target
        node = outline.find(target)
        if node is None:
            raise ValueError(
                f"no heading matches {target!r} in this document"
            )
        return node

    start_node = _resolve(start)
    end_paragraph_index: int
    if end is None:
        end_paragraph_index = len(document.paragraphs)
    else:
        end_node = _resolve(end)
        end_paragraph_index = end_node.paragraph_index

    if end_paragraph_index < start_node.paragraph_index:
        raise ValueError(
            "end heading precedes start heading in document order"
        )

    new_doc = _DocumentFactory()
    # -- the bundled default template ships with one empty paragraph; drop
    # -- it so the slice begins cleanly with the start heading. --
    body = new_doc._element.body  # type: ignore[attr-defined]
    for p in list(body.xpath("./w:p")):
        body.remove(p)

    paragraphs = list(document.paragraphs)
    for p in paragraphs[start_node.paragraph_index:end_paragraph_index]:
        new_doc.append_paragraph(p)

    return new_doc


__all__: "Sequence[str]" = (
    "Outline",
    "OutlineNode",
    "build_outline",
    "slice_document",
)
