"""Executive-summary helpers — Amazon-style narrative documents.

Closes #296.

This module provides two helpers that build Amazon-style narrative
documents in one call::

    from docx import Document
    from docx.kit import exec_summary

    doc = Document()
    exec_summary.one_pager(
        doc,
        title="Project Frobnitz: Q3 update",
        purpose="Decide whether to ship FrobnitzPro by end of Q3.",
        background="Customers have asked for...",
        current_state="Engineering is 80% done...",
        proposal="Ship in two waves...",
        risks=["Dependency on team B", "Holiday freeze cuts review window"],
        asks=["Approval to ship", "Reviewer time on Mar 5"],
    )
    exec_summary.six_pager(
        doc,
        title="FrobnitzPro launch plan",
        sections={
            "Background":            "...",
            "Goals":                 "...",
            "Tenets":                ["Customer obsession", "Speed"],
            "State of the business": "...",
            "Lessons learned":       "...",
            "Strategic priorities":  ["Foundation", "Growth", "Trust"],
            "Looking forward":       "...",
        },
    )
    doc.save("exec.docx")

Both helpers append their content at the end of the document body and
return the list of newly-appended :class:`Paragraph` objects in
document order (including the trailing page-break paragraph when
``page_break=True``, the default). The contract matches the rest of
:mod:`docx.kit` — compose-only against the public python-docx surface,
no XML reach-down, fall back to ``Normal`` when a built-in style is
missing.

The 1-pager follows Amazon's canonical "narrative-on-a-page" shape:
**Purpose / Background / Current state / Proposal / Risks / Asks**.
Risks and Asks accept either a single string (rendered as a single
paragraph) or a sequence of strings (rendered as a bulleted list) so
callers can mix prose and explicit enumerations to taste.

The 6-pager is more flexible — the caller supplies an ordered
``sections`` dict of ``{heading: body}`` pairs, and the helper renders
each as ``Heading 2`` plus body paragraphs. Section order is
preserved (Python 3.7+ dict ordering). Bodies that are sequences
become bulleted lists; string bodies are rendered as one paragraph
each (split on blank lines for multi-paragraph prose).

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Mapping, Optional, Sequence, Union

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph


# -- Word built-in styles the helpers prefer; fall back to "Normal" when
# -- a caller-supplied template has stripped them. --
_STYLE_HEADING_1 = "Heading 1"
_STYLE_HEADING_2 = "Heading 2"
_STYLE_LIST_BULLET = "List Bullet"
_STYLE_NORMAL = "Normal"


def _has_style(document: "Document", style_name: str) -> bool:
    """Return |True| when `document` defines a style named `style_name`."""
    try:
        styles = document.styles
    except Exception:  # pragma: no cover - defensive
        return False
    try:
        styles[style_name]
        return True
    except KeyError:
        return False


def _resolve_style(document: "Document", preferred: str) -> str:
    """Return `preferred` if defined on `document`, else ``"Normal"``."""
    return preferred if _has_style(document, preferred) else _STYLE_NORMAL


def _coerce_paragraphs(body: str) -> List[str]:
    """Split `body` on blank lines into a list of non-empty paragraph strings."""
    chunks = [chunk.strip() for chunk in body.split("\n\n")]
    return [chunk for chunk in chunks if chunk]


def _add_bulleted_list(
    document: "Document", items: Sequence[str]
) -> List["Paragraph"]:
    """Append each item as a ``List Bullet``-styled paragraph (or fallback)."""
    style = _resolve_style(document, _STYLE_LIST_BULLET)
    paragraphs: List["Paragraph"] = []
    for item in items:
        if item is None:
            continue
        text = str(item).strip()
        if not text:
            continue
        paragraphs.append(document.add_paragraph(text, style=style))
    return paragraphs


def _add_section(
    document: "Document",
    heading: str,
    body: Union[str, Sequence[str]],
    heading_level: int,
) -> List["Paragraph"]:
    """Append a heading plus body content; sequences become bulleted lists."""
    paragraphs: List["Paragraph"] = []
    paragraphs.append(document.add_heading(heading, level=heading_level))
    if isinstance(body, str):
        for chunk in _coerce_paragraphs(body):
            paragraphs.append(document.add_paragraph(chunk))
    else:
        # -- Sequence -> bulleted list. --
        paragraphs.extend(_add_bulleted_list(document, list(body)))
    return paragraphs


def _normalise_listish(
    value: Union[str, Sequence[str]],
) -> Union[str, List[str]]:
    """Pass a string through; convert any other sequence into a list of str."""
    if isinstance(value, str):
        return value
    return [str(item) for item in value]


# -- 1-pager ---------------------------------------------------------------


def one_pager(
    document: "Document",
    *,
    title: str,
    purpose: str,
    background: str,
    current_state: str,
    proposal: str,
    risks: Union[str, Sequence[str]],
    asks: Union[str, Sequence[str]],
    page_break: bool = True,
) -> List["Paragraph"]:
    """Append an Amazon-style 1-pager narrative to `document`.

    Renders the canonical six-section "narrative-on-a-page" structure
    Amazon executives use for decision documents:

    1. ``title`` — rendered as ``Heading 1``.
    2. **Purpose** — single paragraph stating the decision being asked.
    3. **Background** — context paragraphs (split on blank lines).
    4. **Current state** — what is true today.
    5. **Proposal** — recommended path forward.
    6. **Risks** — single paragraph or bulleted list of risk lines.
    7. **Asks** — single paragraph or bulleted list of explicit asks.

    Each section heading is rendered as ``Heading 2`` so a reader can
    scan the document outline / TOC to land on any section. The body
    of each section accepts a string (rendered as one or more
    paragraphs, split on blank lines) for the prose sections; ``risks``
    and ``asks`` additionally accept a sequence of strings, in which
    case each entry becomes a ``List Bullet`` paragraph.

    Parameters
    ----------
    document
        The :class:`Document` to mutate; sections are appended at the
        end of the body.
    title
        The 1-pager title. Required — rendered as ``Heading 1``.
    purpose
        One-sentence statement of what decision is being requested.
    background
        Context paragraphs. Multi-paragraph prose may be passed as a
        single string with blank-line separators.
    current_state
        What is true today, before the proposed change.
    proposal
        The recommended path forward.
    risks
        Risk text. Pass a string for a single paragraph, or a sequence
        of strings to render each as a bullet point.
    asks
        Explicit asks of the reader (approvals, reviewer time, …). Pass
        a string for a single paragraph, or a sequence of strings to
        render each as a bullet point.
    page_break
        When |True| (the default), append a trailing page break so the
        next content lands on a fresh page. Pass |False| to suppress.

    Returns
    -------
    list[Paragraph]
        The list of newly-appended paragraphs in document order,
        including the trailing page-break paragraph when emitted.

    Raises
    ------
    ValueError
        When ``title`` is empty, or when any of the required prose
        sections (``purpose`` / ``background`` / ``current_state`` /
        ``proposal``) is empty.

    .. versionadded:: 2026.05.29
    """
    if not title or not title.strip():
        raise ValueError("title must be a non-empty string")
    for name, value in (
        ("purpose", purpose),
        ("background", background),
        ("current_state", current_state),
        ("proposal", proposal),
    ):
        if not value or not str(value).strip():
            raise ValueError("%s must be a non-empty string" % name)

    paragraphs: List["Paragraph"] = []

    # -- Title (Heading 1) --
    paragraphs.append(document.add_heading(title, level=1))

    # -- Prose sections (Heading 2 + body paragraphs) --
    paragraphs.extend(_add_section(document, "Purpose", purpose, heading_level=2))
    paragraphs.extend(
        _add_section(document, "Background", background, heading_level=2)
    )
    paragraphs.extend(
        _add_section(document, "Current state", current_state, heading_level=2)
    )
    paragraphs.extend(
        _add_section(document, "Proposal", proposal, heading_level=2)
    )

    # -- Risks / Asks accept str OR list[str]: --
    paragraphs.extend(
        _add_section(document, "Risks", _normalise_listish(risks), heading_level=2)
    )
    paragraphs.extend(
        _add_section(document, "Asks", _normalise_listish(asks), heading_level=2)
    )

    if page_break:
        paragraphs.append(document.add_page_break())

    return paragraphs


# -- 6-pager ---------------------------------------------------------------


def six_pager(
    document: "Document",
    *,
    title: str,
    sections: Mapping[str, Union[str, Sequence[str]]],
    page_break: bool = True,
) -> List["Paragraph"]:
    """Append an Amazon-style 6-pager deep narrative to `document`.

    Renders the title as ``Heading 1`` followed by one ``Heading 2``
    section per entry in ``sections``. Section order is preserved
    (``sections`` is iterated in insertion order — Python 3.7+ dicts
    are ordered). Each section's body may be:

    * A string — rendered as one paragraph (or several when the string
      contains blank-line separators ``"\\n\\n"``).
    * A sequence of strings — rendered as a bulleted list (one
      ``List Bullet`` paragraph per item).

    The Amazon canonical 6-pager structure (Background / Goals /
    Tenets / State of the business / Lessons learned / Strategic
    priorities / Looking forward) is *not* enforced — the helper is a
    flexible scaffold so callers can structure deep narratives that
    deviate from the canonical seven sections. Pass the canonical
    headings explicitly as the ``sections`` keys to match it.

    Parameters
    ----------
    document
        The :class:`Document` to mutate; the 6-pager is appended at the
        end of the body.
    title
        The 6-pager title. Required — rendered as ``Heading 1``.
    sections
        Ordered mapping of ``{heading: body}`` pairs. ``heading`` is
        rendered as ``Heading 2``; ``body`` is a string (one or more
        paragraphs) or a sequence of strings (a bulleted list). Empty
        bodies render only the heading. At least one section is
        required.
    page_break
        When |True| (the default), append a trailing page break so the
        next content lands on a fresh page. Pass |False| to suppress.

    Returns
    -------
    list[Paragraph]
        The list of newly-appended paragraphs in document order,
        including the trailing page-break paragraph when emitted.

    Raises
    ------
    ValueError
        When ``title`` is empty, when ``sections`` is empty, or when
        any section heading is empty / blank.

    .. versionadded:: 2026.05.29
    """
    if not title or not title.strip():
        raise ValueError("title must be a non-empty string")
    if not sections:
        raise ValueError("sections must contain at least one entry")

    # -- Validate every heading up front so we don't half-emit a
    # -- broken 6-pager before raising. --
    for heading in sections.keys():
        if not heading or not str(heading).strip():
            raise ValueError("every section heading must be a non-empty string")

    paragraphs: List["Paragraph"] = []

    # -- Title --
    paragraphs.append(document.add_heading(title, level=1))

    # -- Sections in caller-supplied order --
    for heading, body in sections.items():
        if body is None:
            paragraphs.append(document.add_heading(str(heading), level=2))
            continue
        if isinstance(body, str):
            paragraphs.extend(
                _add_section(document, str(heading), body, heading_level=2)
            )
        else:
            paragraphs.extend(
                _add_section(
                    document,
                    str(heading),
                    [str(item) for item in body],
                    heading_level=2,
                )
            )

    if page_break:
        paragraphs.append(document.add_page_break())

    return paragraphs


__all__ = [
    "one_pager",
    "six_pager",
]
