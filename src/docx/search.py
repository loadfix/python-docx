"""Search and replace functionality for python-docx documents."""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, List, Tuple

if TYPE_CHECKING:
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run


class SearchMatch:
    """A single match of a search term within a document.

    Provides access to the paragraph containing the match, the run indices that span the
    match, and the character offsets within the reconstructed paragraph text.
    """

    def __init__(
        self,
        paragraph: Paragraph,
        paragraph_index: int,
        run_indices: List[int],
        start: int,
        end: int,
    ):
        self._paragraph = paragraph
        self._paragraph_index = paragraph_index
        self._run_indices = run_indices
        self._start = start
        self._end = end

    @property
    def paragraph(self) -> Paragraph:
        """The |Paragraph| containing this match."""
        return self._paragraph

    @property
    def paragraph_index(self) -> int:
        """Index of the paragraph in the document's paragraph list."""
        return self._paragraph_index

    @property
    def run_indices(self) -> List[int]:
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


def _build_char_map(runs: List[Run]) -> Tuple[str, List[Tuple[int, int]]]:
    """Build full text from runs and a map from character position to (run_index, offset).

    Returns a tuple of (full_text, char_map) where char_map[i] is (run_index,
    char_offset_within_run) for the i-th character in full_text.
    """
    full_text = ""
    char_map: List[Tuple[int, int]] = []
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
    paragraphs: List[Paragraph],
    text: str,
    case_sensitive: bool = True,
    whole_word: bool = False,
) -> List[SearchMatch]:
    """Find all occurrences of `text` across `paragraphs`.

    Returns a list of |SearchMatch| objects, one for each occurrence found.
    """
    if not text:
        return []

    pattern = _compile_pattern(text, case_sensitive, whole_word)
    matches: List[SearchMatch] = []

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
    paragraphs: List[Paragraph],
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
    runs: List[Run],
    char_map: List[Tuple[int, int]],
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
