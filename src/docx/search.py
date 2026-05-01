"""Search and replace functionality for python-docx documents."""

from __future__ import annotations

import re
from collections.abc import Iterator
from typing import TYPE_CHECKING, Union

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run

RegexPattern = Union[str, re.Pattern[str]]


class SearchMatch:
    """A single match of a search term within a document.

    Provides access to the paragraph containing the match, the run indices that span the
    match, and the character offsets within the reconstructed paragraph text.

    When the match comes from a cross-story search (see :meth:`Document.search_all`),
    :attr:`location` identifies which "story" the paragraph belongs to — for example
    ``"body"``, ``"table:0:row:1:col:2"``, ``"header:section0:primary"``,
    ``"footnote:2"``, ``"endnote:3"``, or ``"comment:5"``. Matches produced by the
    body-only helpers carry :attr:`location` of |None|.
    """

    def __init__(
        self,
        paragraph: Paragraph,
        paragraph_index: int,
        run_indices: list[int],
        start: int,
        end: int,
        location: str | None = None,
    ):
        self._paragraph = paragraph
        self._paragraph_index = paragraph_index
        self._run_indices = run_indices
        self._start = start
        self._end = end
        self._location = location

    @property
    def paragraph(self) -> Paragraph:
        """The |Paragraph| containing this match."""
        return self._paragraph

    @property
    def paragraph_index(self) -> int:
        """Index of the paragraph in its story's paragraph list.

        For a cross-story match, the index is relative to the paragraphs of the
        specific story identified by :attr:`location` (e.g. the paragraphs of a
        particular footer or footnote), not a global document-wide index.
        """
        return self._paragraph_index

    @property
    def run_indices(self) -> list[int]:
        """Indices of runs that span this match."""
        return self._run_indices

    @property
    def start(self) -> int:
        """Character offset of match start in the paragraph's reconstructed text."""
        return self._start

    @property
    def end(self) -> int:
        """Character offset of match end in the paragraph's reconstructed text."""
        return self._end

    @property
    def location(self) -> str | None:
        """Story-identifier for this match, or |None| if unknown.

        Body-only search helpers leave this |None|; :meth:`Document.search_all` and
        friends populate this with a string like ``"body"``, ``"table:0:row:1:col:2"``,
        ``"header:section0:primary"``, ``"footnote:2"``, ``"endnote:3"``, or
        ``"comment:5"``.
        """
        return self._location


def _build_char_map(runs: list[Run]) -> tuple[str, list[tuple[int, int]]]:
    """Build full text from runs and a map from character position to (run_index, offset).

    Returns a tuple of (full_text, char_map) where char_map[i] is (run_index,
    char_offset_within_run) for the i-th character in full_text.
    """
    full_text = ""
    char_map: list[tuple[int, int]] = []
    for run_idx, run in enumerate(runs):
        run_text = run.text
        for char_offset in range(len(run_text)):
            char_map.append((run_idx, char_offset))
        full_text += run_text
    return full_text, char_map


def _compile_pattern(text: str, case_sensitive: bool, whole_word: bool) -> re.Pattern[str]:
    """Compile a regex pattern for the given search text and options."""
    escaped = re.escape(text)
    if whole_word:
        escaped = rf"\b{escaped}\b"
    flags = 0 if case_sensitive else re.IGNORECASE
    return re.compile(escaped, flags)


def search_paragraphs(
    paragraphs: list[Paragraph],
    text: str,
    case_sensitive: bool = True,
    whole_word: bool = False,
) -> list[SearchMatch]:
    """Find all occurrences of `text` across `paragraphs`.

    Returns a list of |SearchMatch| objects, one for each occurrence found.
    """
    if not text:
        return []

    pattern = _compile_pattern(text, case_sensitive, whole_word)
    matches: list[SearchMatch] = []

    for para_idx, paragraph in enumerate(paragraphs):
        full_text, char_map = _build_char_map(paragraph.runs)
        for m in pattern.finditer(full_text):
            start, end = m.start(), m.end()
            run_indices = sorted({char_map[i][0] for i in range(start, end)})
            matches.append(
                SearchMatch(
                    paragraph=paragraph,
                    paragraph_index=para_idx,
                    run_indices=run_indices,
                    start=start,
                    end=end,
                )
            )

    return matches


def replace_in_paragraphs(
    paragraphs: list[Paragraph],
    old_text: str,
    new_text: str,
    case_sensitive: bool = True,
    whole_word: bool = False,
) -> int:
    """Replace all occurrences of `old_text` with `new_text` in `paragraphs`.

    Preserves the formatting of the first character's run for each replacement. Returns
    the number of replacements made.
    """
    if not old_text:
        return 0

    pattern = _compile_pattern(old_text, case_sensitive, whole_word)
    total_replacements = 0

    for paragraph in paragraphs:
        total_replacements += _replace_in_paragraph(paragraph, pattern, new_text)

    return total_replacements


def _replace_in_paragraph(
    paragraph: Paragraph, pattern: re.Pattern[str], new_text: str
) -> int:
    """Replace all matches of `pattern` with `new_text` in a single paragraph.

    Processes matches from right to left so that earlier character positions remain valid
    as the text is modified.
    """
    runs = paragraph.runs
    if not runs:
        return 0

    full_text, char_map = _build_char_map(runs)
    matches = list(pattern.finditer(full_text))
    if not matches:
        return 0

    # Process matches from right to left to preserve positions.
    for m in reversed(matches):
        _apply_replacement(runs, char_map, m.start(), m.end(), new_text)

    return len(matches)


def _apply_replacement(
    runs: list[Run],
    char_map: list[tuple[int, int]],
    match_start: int,
    match_end: int,
    new_text: str,
) -> None:
    """Replace the text at [match_start, match_end) with `new_text` across runs.

    The formatting of the run containing the first matched character is preserved. Text
    is removed from subsequent runs that were part of the match; empty runs are left in
    place (their formatting may be needed by Word).
    """
    first_run_idx, first_char_offset = char_map[match_start]
    last_run_idx, last_char_offset = char_map[match_end - 1]

    first_run = runs[first_run_idx]
    first_run_text = first_run.text

    if first_run_idx == last_run_idx:
        # Match is entirely within one run.
        first_run.text = (
            first_run_text[:first_char_offset]
            + new_text
            + first_run_text[last_char_offset + 1 :]
        )
    else:
        # Match spans multiple runs. Put replacement text in the first run,
        # clear matched portions from the remaining runs.
        first_run.text = first_run_text[:first_char_offset] + new_text

        # Clear text from fully-spanned middle runs.
        for run_idx in range(first_run_idx + 1, last_run_idx):
            runs[run_idx].text = ""

        # Trim the matched prefix from the last run.
        last_run = runs[last_run_idx]
        last_run.text = last_run.text[last_char_offset + 1 :]


def _coerce_regex(pattern: RegexPattern, flags: int = 0) -> re.Pattern[str]:
    """Return a compiled regex pattern.

    If `pattern` is already a compiled `re.Pattern`, it is returned unchanged and
    `flags` is ignored. Otherwise `pattern` is compiled with `flags`.
    """
    if isinstance(pattern, re.Pattern):
        return pattern
    return re.compile(pattern, flags)


def search_paragraphs_regex(
    paragraphs: list[Paragraph],
    pattern: RegexPattern,
    flags: int = 0,
) -> list[SearchMatch]:
    """Find all regex matches of `pattern` across `paragraphs`.

    `pattern` may be a string or a compiled `re.Pattern`. When `pattern` is a string,
    `flags` (e.g. `re.IGNORECASE`) is applied when compiling. When `pattern` is already
    compiled, `flags` is ignored. Returns a list of |SearchMatch| objects, one for each
    match found.
    """
    compiled = _coerce_regex(pattern, flags)
    matches: list[SearchMatch] = []

    for para_idx, paragraph in enumerate(paragraphs):
        full_text, char_map = _build_char_map(paragraph.runs)
        for m in compiled.finditer(full_text):
            start, end = m.start(), m.end()
            # For zero-width matches, run_indices is empty since no characters are
            # spanned. Otherwise, collect unique run indices covering [start, end).
            run_indices = sorted({char_map[i][0] for i in range(start, end)})
            matches.append(
                SearchMatch(
                    paragraph=paragraph,
                    paragraph_index=para_idx,
                    run_indices=run_indices,
                    start=start,
                    end=end,
                )
            )

    return matches


def replace_in_paragraphs_regex(
    paragraphs: list[Paragraph],
    pattern: RegexPattern,
    replacement: str,
    flags: int = 0,
) -> int:
    """Replace all regex matches of `pattern` with `replacement` in `paragraphs`.

    `pattern` may be a string or a compiled `re.Pattern`. When `pattern` is a string,
    `flags` (e.g. `re.IGNORECASE`) is applied when compiling. `replacement` follows
    `re.sub` semantics — backreferences such as ``\\1`` and ``\\g<name>`` are expanded
    per match. Preserves the formatting of the first character's run for each
    replacement. Returns the number of replacements made.
    """
    compiled = _coerce_regex(pattern, flags)
    total_replacements = 0

    for paragraph in paragraphs:
        total_replacements += _replace_in_paragraph_regex(
            paragraph, compiled, replacement
        )

    return total_replacements


def _replace_in_paragraph_regex(
    paragraph: Paragraph, pattern: re.Pattern[str], replacement: str
) -> int:
    """Replace all regex matches of `pattern` with `replacement` in one paragraph.

    Each match's replacement text is produced via `Match.expand()` so that backreferences
    are resolved. Matches are applied right-to-left so earlier character positions remain
    valid as the text is modified.
    """
    runs = paragraph.runs
    if not runs:
        return 0

    full_text, char_map = _build_char_map(runs)
    # Skip zero-width matches — they have no characters to replace and can't be
    # positioned within runs unambiguously.
    matches = [m for m in pattern.finditer(full_text) if m.end() > m.start()]
    if not matches:
        return 0

    for m in reversed(matches):
        expanded = m.expand(replacement)
        _apply_replacement(runs, char_map, m.start(), m.end(), expanded)

    return len(matches)


# -- Cross-story search / replace --------------------------------------------


def _iter_all_paragraphs(
    document: Document,
) -> Iterator[tuple[list[Paragraph], str]]:
    """Yield ``(paragraphs, location)`` pairs covering every story in `document`.

    Stories visited, in order:

    - Body paragraphs, tagged ``"body"``.
    - Paragraphs in body-level tables, tagged ``"table:<t>:row:<r>:col:<c>"``.
      Only top-level body tables are descended into; tables nested inside header,
      footer, footnote, endnote, or comment stories are not visited, and tables
      nested inside body tables are likewise skipped (cells already provide the
      searchable text for their immediate contents).
    - Headers and footers for each section, one pair per non-linked definition,
      tagged ``"header:section<i>:primary"`` (or ``even_page``/``first_page``)
      and ``"footer:...":`` likewise. Sections that simply inherit the previous
      section's header/footer are not re-emitted.
    - Footnote paragraphs, tagged ``"footnote:<id>"``.
    - Endnote paragraphs, tagged ``"endnote:<id>"``.
    - Comment paragraphs, tagged ``"comment:<id>"``.
    """
    # -- body --
    yield list(document.paragraphs), "body"

    # -- body tables (top-level only; no recursion into nested tables) --
    for t_idx, table in enumerate(document.tables):
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                yield (
                    list(cell.paragraphs),
                    f"table:{t_idx}:row:{r_idx}:col:{c_idx}",
                )

    # -- headers / footers (skip inherited / linked-to-previous definitions) --
    _hf_kinds: tuple[tuple[str, str], ...] = (
        ("header", "primary"),
        ("header", "even_page"),
        ("header", "first_page"),
        ("footer", "primary"),
        ("footer", "even_page"),
        ("footer", "first_page"),
    )
    for s_idx, section in enumerate(document.sections):
        accessors = {
            ("header", "primary"): section.header,
            ("header", "even_page"): section.even_page_header,
            ("header", "first_page"): section.first_page_header,
            ("footer", "primary"): section.footer,
            ("footer", "even_page"): section.even_page_footer,
            ("footer", "first_page"): section.first_page_footer,
        }
        for kind, variant in _hf_kinds:
            hf = accessors[(kind, variant)]
            if hf.is_linked_to_previous:
                continue
            yield (
                list(hf.paragraphs),
                f"{kind}:section{s_idx}:{variant}",
            )

    # -- footnotes / endnotes / comments --
    # These attributes create default parts lazily; guard against exotic Document
    # configurations (e.g. unit-test fixtures built around a Mock part) where the
    # accessors raise instead of returning an iterable.
    try:
        footnotes_iter = list(document.footnotes)
    except (AttributeError, KeyError, AssertionError, TypeError):
        footnotes_iter = []
    for footnote in footnotes_iter:
        yield list(footnote.paragraphs), f"footnote:{footnote.footnote_id}"

    try:
        endnotes_iter = list(document.endnotes)
    except (AttributeError, KeyError, AssertionError, TypeError):
        endnotes_iter = []
    for endnote in endnotes_iter:
        yield list(endnote.paragraphs), f"endnote:{endnote.endnote_id}"

    try:
        comments_iter = list(document.comments)
    except (AttributeError, KeyError, AssertionError, TypeError):
        comments_iter = []
    for comment in comments_iter:
        yield list(comment.paragraphs), f"comment:{comment.comment_id}"


def _search_in_story(
    paragraphs: list[Paragraph],
    pattern: re.Pattern[str],
    location: str,
) -> list[SearchMatch]:
    """Run a compiled `pattern` across `paragraphs`, tagging each match with `location`."""
    matches: list[SearchMatch] = []
    for para_idx, paragraph in enumerate(paragraphs):
        full_text, char_map = _build_char_map(paragraph.runs)
        for m in pattern.finditer(full_text):
            start, end = m.start(), m.end()
            run_indices = sorted({char_map[i][0] for i in range(start, end)})
            matches.append(
                SearchMatch(
                    paragraph=paragraph,
                    paragraph_index=para_idx,
                    run_indices=run_indices,
                    start=start,
                    end=end,
                    location=location,
                )
            )
    return matches


def search_all_paragraphs(
    document: Document,
    text: str,
    case_sensitive: bool = True,
    whole_word: bool = False,
) -> list[SearchMatch]:
    """Find all occurrences of `text` across every story in `document`.

    Returns a list of |SearchMatch| objects with :attr:`SearchMatch.location` set to
    indicate which story each match came from. The stories searched include the
    body, body-level tables, non-inherited headers and footers on every section,
    footnotes, endnotes, and comments.
    """
    if not text:
        return []

    pattern = _compile_pattern(text, case_sensitive, whole_word)
    matches: list[SearchMatch] = []
    for paragraphs, location in _iter_all_paragraphs(document):
        matches.extend(_search_in_story(paragraphs, pattern, location))
    return matches


def search_all_paragraphs_regex(
    document: Document,
    pattern: RegexPattern,
    flags: int = 0,
) -> list[SearchMatch]:
    """Find all regex matches of `pattern` across every story in `document`.

    Returns a list of |SearchMatch| objects with :attr:`SearchMatch.location` populated.
    Stories searched are the same as for :func:`search_all_paragraphs`.
    """
    compiled = _coerce_regex(pattern, flags)
    matches: list[SearchMatch] = []
    for paragraphs, location in _iter_all_paragraphs(document):
        matches.extend(_search_in_story(paragraphs, compiled, location))
    return matches


def replace_in_all_paragraphs(
    document: Document,
    old_text: str,
    new_text: str,
    case_sensitive: bool = True,
    whole_word: bool = False,
) -> int:
    """Replace `old_text` with `new_text` in every story of `document`.

    Stories updated are the same as those searched by :func:`search_all_paragraphs`.
    Returns the total number of replacements made across all stories.
    """
    if not old_text:
        return 0

    pattern = _compile_pattern(old_text, case_sensitive, whole_word)
    total = 0
    for paragraphs, _ in _iter_all_paragraphs(document):
        for paragraph in paragraphs:
            total += _replace_in_paragraph(paragraph, pattern, new_text)
    return total


def replace_in_all_paragraphs_regex(
    document: Document,
    pattern: RegexPattern,
    replacement: str,
    flags: int = 0,
) -> int:
    """Replace regex matches of `pattern` with `replacement` across every story.

    Stories updated are the same as those searched by :func:`search_all_paragraphs_regex`.
    Returns the total number of replacements made.
    """
    compiled = _coerce_regex(pattern, flags)
    total = 0
    for paragraphs, _ in _iter_all_paragraphs(document):
        for paragraph in paragraphs:
            total += _replace_in_paragraph_regex(paragraph, compiled, replacement)
    return total
