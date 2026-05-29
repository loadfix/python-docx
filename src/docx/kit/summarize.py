"""Token-bounded progressive document summarisation.

Closes #303.

Long-form documents (reports, design docs, RFPs, theses) are routinely
fed into LLM context windows or skim-reading workflows where the
*entire* body is too long to consume but a *fixed-budget* summary is
useful. The :func:`summarize` helper walks a |Document|, splits the
body into sections (by ``Heading 1``, falling back to ``Heading 2``
when no H1 is present, falling back to chunks of ``N`` paragraphs when
no headings are present), allocates a token budget across sections
proportional to each section's length, and emits a per-section summary
under that budget.

Two output shapes::

    from docx import Document
    from docx.kit import summarize

    doc = Document("long_report.docx")

    # 1. Structured: list[dict] with section / summary / tokens keys.
    sections = summarize.summarize(doc, max_tokens=500)
    # [
    #   {"section": "Introduction",  "summary": "...", "tokens": 80},
    #   {"section": "Methods",       "summary": "...", "tokens": 120},
    #   ...
    # ]

    # 2. Flat string with bold section headers.
    text = summarize.as_text(doc, max_tokens=500)

The default :func:`summariser` is *structural-extractive*: it splits
the section body into sentences and emits the first ``N`` sentences
that fit in the section's allocated budget. Callers who want LLM-grade
summaries inject a custom callable::

    def my_llm_summarise(text, max_tokens):
        return openai_summarise(text, max_tokens=max_tokens)

    sections = summarize.summarize(
        doc, max_tokens=500, summariser=my_llm_summarise
    )

The token counter is approximate: a 4-chars-per-token heuristic is
used by default (close enough for OpenAI / Anthropic models for budget
allocation; not for billing). Callers who need a real tokenizer
(``tiktoken``, ``transformers``) inject one via ``token_counter=``::

    import tiktoken
    enc = tiktoken.get_encoding("cl100k_base")

    sections = summarize.summarize(
        doc,
        max_tokens=500,
        token_counter=lambda s: len(enc.encode(s)),
    )

The helper composes only python-docx's *public* API (iterates
``Document.paragraphs`` and reads ``Paragraph.text`` / ``Paragraph.style``).
No XML reach-down. No PyPI dependencies are introduced — ``tiktoken``
/ ``openai`` / ``anthropic`` are reachable only through the
``token_counter=`` / ``summariser=`` injection points.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

import re
from typing import (
    TYPE_CHECKING,
    Callable,
    Dict,
    List,
    Optional,
    Sequence,
    Union,
)

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph


# ---------------------------------------------------------------------------
# Public types

# A section dict has three keys: ``section`` (str), ``summary`` (str),
# ``tokens`` (int). Aliased loosely for readability — a TypedDict would
# inflate the import surface for an internal data shape callers will
# treat as a plain dict anyway.
SectionSummary = Dict[str, Union[str, int]]

# Pluggable callbacks.
TokenCounter = Callable[[str], int]
Summariser = Callable[[str, int], str]
SectionPredicate = Callable[["Paragraph"], bool]


# ---------------------------------------------------------------------------
# Tunable defaults

# When a document has no headings, this is the chunk size used to
# split the body. Twenty paragraphs is roughly two screens of body
# text — long enough that each chunk has summarisable substance,
# short enough that a 30-paragraph essay produces two summary rows.
_DEFAULT_HEADINGLESS_CHUNK = 20

# Floor on per-section budget. Anything below this rounds to zero
# tokens which the summariser collapses to an empty string. A
# 10-token floor keeps short sections at least a fragment long.
_MIN_SECTION_BUDGET = 10

# Sentence splitter — cheap regex over ``.``, ``!``, ``?`` followed by
# whitespace. Not perfect (won't survive ``Dr. Smith``) but good enough
# for an extractive baseline; callers wanting better fidelity inject a
# real summariser.
_SENTENCE_SPLIT = re.compile(r"(?<=[.!?])\s+")


# ---------------------------------------------------------------------------
# Token counting


def count_tokens(text: str) -> int:
    """Return an approximate token count for ``text``.

    Uses a 4-chars-per-token heuristic — close enough for budget
    allocation against OpenAI / Anthropic / Google models (their real
    tokenisers vary by ~10-20% on English prose). Not suitable for
    billing or hard context-window enforcement; for those callers
    inject a real tokeniser via ``token_counter=`` on
    :func:`summarize` / :func:`as_text`.

    Returns 0 for empty input.
    """
    if not text:
        return 0
    # -- Round up so a 5-character string scores 2 tokens, not 1.
    # -- This avoids the degenerate case where a 3-character section
    # -- gets a 0-token budget and is dropped entirely.
    return max(1, (len(text) + 3) // 4)


# ---------------------------------------------------------------------------
# Section detection


def _is_heading(paragraph: "Paragraph", level: int) -> bool:
    """Return True when ``paragraph`` carries the ``Heading <level>`` style."""
    style = paragraph.style
    if style is None:
        return False
    name = getattr(style, "name", None) or ""
    return name == f"Heading {level}"


def _detect_heading_level(document: "Document") -> Optional[int]:
    """Return ``1`` if the document has any ``Heading 1``, else ``2`` if it
    has any ``Heading 2``, else ``None``.

    The fallback chain mirrors how human readers parse documents:
    primary structure first (chapters), then secondary (sections),
    then "give up and chunk".
    """
    for paragraph in document.paragraphs:
        if _is_heading(paragraph, 1):
            return 1
    for paragraph in document.paragraphs:
        if _is_heading(paragraph, 2):
            return 2
    return None


def _split_by_heading(
    document: "Document",
    level: int,
    section_predicate: Optional[SectionPredicate],
) -> List[Dict[str, object]]:
    """Split ``document.paragraphs`` into sections demarcated by headings.

    Body paragraphs that precede the first heading are collected into
    a leading "Section 1" row so a doc that opens with body text and
    then introduces headings doesn't lose its preamble.

    When ``section_predicate`` is supplied the heading test routes
    through it instead of the default style-name check (handy for
    callers whose documents use custom heading styles).
    """
    sections: List[Dict[str, object]] = []
    current_title: Optional[str] = None
    current_body: List[str] = []

    def _is_section_header(paragraph: "Paragraph") -> bool:
        if section_predicate is not None:
            return bool(section_predicate(paragraph))
        return _is_heading(paragraph, level)

    def _flush() -> None:
        # -- Drop sections with no title *and* no body — those are
        # -- empty preambles. Sections with a title but no body still
        # -- emit (the heading itself was a meaningful split point).
        if current_title is None and not current_body:
            return
        sections.append(
            {
                "section": current_title or f"Section {len(sections) + 1}",
                "body": "\n".join(current_body).strip(),
            }
        )

    for paragraph in document.paragraphs:
        text = paragraph.text or ""
        if _is_section_header(paragraph):
            _flush()
            current_title = text.strip() or f"Section {len(sections) + 1}"
            current_body = []
        else:
            if text.strip():
                current_body.append(text)

    _flush()
    return sections


def _split_by_chunks(
    document: "Document", chunk_size: int
) -> List[Dict[str, object]]:
    """Split ``document.paragraphs`` into ``chunk_size``-paragraph chunks.

    Used when the document has no headings. Empty paragraphs are
    skipped so a document padded with whitespace doesn't produce
    empty rows.
    """
    bodies = [p.text for p in document.paragraphs if (p.text or "").strip()]
    sections: List[Dict[str, object]] = []
    for index in range(0, len(bodies), chunk_size):
        chunk = bodies[index : index + chunk_size]
        sections.append(
            {
                "section": f"Section {len(sections) + 1}",
                "body": "\n".join(chunk).strip(),
            }
        )
    return sections


# ---------------------------------------------------------------------------
# Default extractive summariser


def _extractive_summarise(
    text: str, max_tokens: int, token_counter: TokenCounter
) -> str:
    """Return the first sentences of ``text`` that fit in ``max_tokens``.

    Splits the body into sentences via :data:`_SENTENCE_SPLIT`, then
    accumulates from the start until the next sentence would push the
    accumulated count over ``max_tokens``. When the *first* sentence
    alone exceeds ``max_tokens``, a character-truncated head is
    emitted so the section never returns an empty string from a
    non-empty body (the caller asked for *something* under the
    budget).
    """
    if not text or max_tokens <= 0:
        return ""

    sentences = [s for s in _SENTENCE_SPLIT.split(text.strip()) if s.strip()]
    if not sentences:
        return ""

    accumulated: List[str] = []
    used = 0
    for sentence in sentences:
        cost = token_counter(sentence)
        if used + cost > max_tokens:
            break
        accumulated.append(sentence)
        used += cost

    if accumulated:
        return " ".join(accumulated).strip()

    # -- The first sentence alone is over budget. Emit a
    # -- character-truncated head so the row carries *something*.
    head_chars = max_tokens * 4
    truncated = sentences[0][:head_chars].rstrip()
    return truncated


# ---------------------------------------------------------------------------
# Budget allocation


def _allocate_budget(
    section_token_lengths: Sequence[int], max_tokens: int
) -> List[int]:
    """Distribute ``max_tokens`` across ``section_token_lengths`` proportionally.

    A section that is 30% of the body gets 30% of the budget, capped
    at the section's own token length (no point allocating 200 tokens
    to a 50-token section). Any rounding remainder is added to the
    largest section so the total stays at most ``max_tokens``.

    Sections shorter than :data:`_MIN_SECTION_BUDGET` get the floor
    (capped at their own length) so a 5-token aside still gets a
    fragment of itself rather than silently dropping out.
    """
    total = sum(section_token_lengths)
    if total <= 0 or max_tokens <= 0:
        return [0] * len(section_token_lengths)

    # -- Proportional cut, capped at the section's own length --
    raw = [
        min(length, max(_MIN_SECTION_BUDGET, (length * max_tokens) // total))
        if length > 0
        else 0
        for length in section_token_lengths
    ]

    # -- The min(length, ...) cap and floor can over- or under-shoot
    # -- the budget. Trim from the largest entries first when we're
    # -- over; nothing to do when we're under (callers asked for at
    # -- *most* max_tokens, not at least).
    over = sum(raw) - max_tokens
    if over > 0:
        # -- Walk the entries in decreasing-budget order, shaving one
        # -- token at a time from each until we're back inside the
        # -- budget. O(over * n) but n is small (one row per section).
        order = sorted(range(len(raw)), key=lambda i: raw[i], reverse=True)
        idx = 0
        while over > 0 and any(b > 0 for b in raw):
            i = order[idx % len(order)]
            if raw[i] > 0:
                raw[i] -= 1
                over -= 1
            idx += 1

    return raw


# ---------------------------------------------------------------------------
# Public API


def summarize(
    document: "Document",
    *,
    max_tokens: int,
    summariser: Optional[Summariser] = None,
    section_predicate: Optional[SectionPredicate] = None,
    token_counter: Optional[TokenCounter] = None,
    chunk_size: int = _DEFAULT_HEADINGLESS_CHUNK,
) -> List[SectionSummary]:
    """Return a list of per-section summaries totalling at most ``max_tokens``.

    Splits ``document`` into sections — by ``Heading 1`` style if
    present, falling back to ``Heading 2`` if not, falling back to
    chunks of ``chunk_size`` paragraphs when the document has no
    headings — then allocates a token budget across sections
    proportional to each section's length and runs ``summariser`` on
    each section's body text under its allocated budget.

    Returns ``list[dict]`` with three keys per row:

    * ``section`` — the heading text (or ``"Section N"`` for unnamed
      chunks);
    * ``summary`` — the section's summary, ``""`` for sections that
      receive a 0-token budget;
    * ``tokens`` — the token count of ``summary`` measured by
      ``token_counter``.

    Returns ``[]`` for an empty document — never raises on empty input.

    :param max_tokens: total token budget across the whole summary.
        Must be a positive integer (``ValueError`` otherwise).
    :param summariser: callable ``(text, max_tokens) -> str``. Defaults
        to a structural-extractive summariser that returns the first
        sentences of each section that fit in the section's budget.
    :param section_predicate: callable ``(paragraph) -> bool`` that
        returns ``True`` when ``paragraph`` is a section header.
        Defaults to a style-name check against the auto-detected
        heading level.
    :param token_counter: callable ``(text) -> int`` returning a token
        count. Defaults to :func:`count_tokens` (4-chars-per-token).
    :param chunk_size: paragraphs-per-chunk for headingless documents.
        Defaults to ``20``.
    """
    if not isinstance(max_tokens, int) or max_tokens <= 0:
        raise ValueError("max_tokens must be a positive integer")
    if chunk_size <= 0:
        raise ValueError("chunk_size must be a positive integer")

    counter: TokenCounter = token_counter or count_tokens

    # -- Section detection. When a section_predicate is supplied we
    # -- skip the heading-level autodetect and route the predicate
    # -- straight through; otherwise auto-detect H1 -> H2 -> chunks.
    if section_predicate is not None:
        sections = _split_by_heading(document, level=1, section_predicate=section_predicate)
    else:
        level = _detect_heading_level(document)
        if level is not None:
            sections = _split_by_heading(document, level=level, section_predicate=None)
        else:
            sections = _split_by_chunks(document, chunk_size=chunk_size)

    # -- Drop empty sections (no title and no body) — defensive; the
    # -- _split_* helpers should already have filtered these.
    sections = [s for s in sections if s.get("section") or s.get("body")]
    if not sections:
        return []

    section_token_lengths = [counter(str(s.get("body", ""))) for s in sections]
    budgets = _allocate_budget(section_token_lengths, max_tokens)

    # -- Resolve the summariser callable. The default closure carries
    # -- the user-supplied token_counter so the extractive path
    # -- measures sentences with the same yardstick as the budget.
    if summariser is not None:
        run = summariser
    else:
        def run(text: str, budget: int) -> str:
            return _extractive_summarise(text, budget, counter)

    out: List[SectionSummary] = []
    for section, budget in zip(sections, budgets):
        body = str(section.get("body", "")).strip()
        if budget <= 0 or not body:
            summary = ""
        else:
            summary = run(body, budget) or ""
        out.append(
            {
                "section": str(section.get("section", "")),
                "summary": summary,
                "tokens": counter(summary),
            }
        )
    return out


def as_text(
    document: "Document",
    *,
    max_tokens: int,
    summariser: Optional[Summariser] = None,
    section_predicate: Optional[SectionPredicate] = None,
    token_counter: Optional[TokenCounter] = None,
    chunk_size: int = _DEFAULT_HEADINGLESS_CHUNK,
) -> str:
    """Return a flat-string summary of ``document`` under ``max_tokens``.

    Calls :func:`summarize` with the same kwargs, then concatenates
    each row into a single string with the section heading rendered as
    a bold-prefixed line followed by the summary body. The bold marker
    is the GitHub-flavoured Markdown ``**heading**`` form so the output
    is human-readable in any plain-text consumer (terminal, email,
    Slack, LLM prompt) and renders as bold in Markdown viewers.

    Empty sections (0-token budget, no body) are omitted from the flat
    text so a 500-token budget on a 30-section doc doesn't render as
    20 empty headings followed by 10 short paragraphs.

    Returns ``""`` for an empty document.
    """
    rows = summarize(
        document,
        max_tokens=max_tokens,
        summariser=summariser,
        section_predicate=section_predicate,
        token_counter=token_counter,
        chunk_size=chunk_size,
    )
    parts: List[str] = []
    for row in rows:
        section = str(row.get("section", "")).strip()
        summary = str(row.get("summary", "")).strip()
        if not summary:
            continue
        if section:
            parts.append(f"**{section}**\n{summary}")
        else:
            parts.append(summary)
    return "\n\n".join(parts)


__all__ = [
    "summarize",
    "as_text",
    "count_tokens",
    "SectionSummary",
    "TokenCounter",
    "Summariser",
    "SectionPredicate",
]
