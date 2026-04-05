"""Document-wide search and replace with formatting preservation."""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, List, NamedTuple, Sequence, Tuple

if TYPE_CHECKING:
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run


class SearchMatch:
    """A single match found by a document search.

    Attributes:
        paragraph: The |Paragraph| containing the match.
        paragraph_index: The index of the paragraph in the document's paragraph list.
        run_indices: A list of run indices (within the paragraph) that contain the match.
        start: The character offset of the match start within the paragraph text.
        end: The character offset of the match end within the paragraph text.
    """

    def __init__(
        self,
        paragraph: Paragraph,
        paragraph_index: int,
        run_indices: List[int],
        start: int,
        end: int,
    ):
        self.paragraph = paragraph
        self.paragraph_index = paragraph_index
        self.run_indices = run_indices
        self.start = start
        self.end = end


class _RunCharMap(NamedTuple):
    """Maps a character offset within paragraph text to a run index and offset within that run."""

    run_index: int
    char_offset: int


def _build_char_map(runs: Sequence[Run]) -> Tuple[str, List[_RunCharMap]]:
    """Build a full-text string and character-to-run mapping for a sequence of runs.

    Returns a tuple of (full_text, char_map) where char_map[i] gives the run index and
    character offset within that run for the i-th character in full_text.
    """
    full_text_parts: List[str] = []
    char_map: List[_RunCharMap] = []
    for run_idx, run in enumerate(runs):
        run_text = run.text
        for char_offset in range(len(run_text)):
            char_map.append(_RunCharMap(run_idx, char_offset))
        full_text_parts.append(run_text)
    return "".join(full_text_parts), char_map


def _find_matches_in_text(
    text: str, search_text: str, case_sensitive: bool, whole_word: bool
) -> List[Tuple[int, int]]:
    """Find all occurrences of search_text in text. Returns list of (start, end) tuples."""
    flags = 0 if case_sensitive else re.IGNORECASE
    pattern = re.escape(search_text)
    if whole_word:
        pattern = r"\b" + pattern + r"\b"
    return [(m.start(), m.end()) for m in re.finditer(pattern, text, flags)]


def _run_indices_for_match(
    char_map: List[_RunCharMap], start: int, end: int
) -> List[int]:
    """Determine which run indices are involved in a match spanning [start, end)."""
    if not char_map or start >= end:
        return []
    indices: List[int] = []
    seen: set[int] = set()
    for i in range(start, end):
        run_idx = char_map[i].run_index
        if run_idx not in seen:
            seen.add(run_idx)
            indices.append(run_idx)
    return indices


def search_paragraphs(
    paragraphs: Sequence[Paragraph],
    search_text: str,
    case_sensitive: bool = True,
    whole_word: bool = False,
) -> List[SearchMatch]:
    """Search all paragraphs for occurrences of `search_text`.

    Returns a list of `SearchMatch` objects, one per match found.
    """
    matches: List[SearchMatch] = []
    for para_idx, paragraph in enumerate(paragraphs):
        runs = paragraph.runs
        if not runs:
            continue
        full_text, char_map = _build_char_map(runs)
        hits = _find_matches_in_text(full_text, search_text, case_sensitive, whole_word)
        for start, end in hits:
            run_indices = _run_indices_for_match(char_map, start, end)
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


def _replace_in_paragraph(
    paragraph: Paragraph,
    old_text: str,
    new_text: str,
    case_sensitive: bool = True,
    whole_word: bool = False,
) -> int:
    """Replace all occurrences of `old_text` with `new_text` in a single paragraph.

    Preserves the formatting of the first character's run for each replacement.
    Returns the number of replacements made.
    """
    runs = paragraph.runs
    if not runs:
        return 0

    full_text, char_map = _build_char_map(runs)
    hits = _find_matches_in_text(full_text, old_text, case_sensitive, whole_word)
    if not hits:
        return 0

    # -- Process replacements in reverse order so earlier offsets remain valid --
    for start, end in reversed(hits):
        _apply_single_replacement(runs, char_map, start, end, new_text)

    return len(hits)


def _apply_single_replacement(
    runs: Sequence[Run],
    char_map: List[_RunCharMap],
    start: int,
    end: int,
    new_text: str,
) -> None:
    """Replace the text at [start, end) across runs with new_text.

    The replacement text is placed in the first affected run, preserving its formatting.
    Text in subsequent affected runs is removed as needed.
    """
    if start >= end or not char_map:
        return

    first_map = char_map[start]
    last_map = char_map[end - 1]

    first_run_idx = first_map.run_index
    first_char_offset = first_map.char_offset
    last_run_idx = last_map.run_index
    last_char_offset = last_map.char_offset

    if first_run_idx == last_run_idx:
        # -- Match is entirely within a single run --
        run = runs[first_run_idx]
        run_text = run.text
        run.text = run_text[:first_char_offset] + new_text + run_text[last_char_offset + 1 :]
    else:
        # -- Match spans multiple runs --
        # 1. Update the first run: keep text before match, append new_text
        first_run = runs[first_run_idx]
        first_run_text = first_run.text
        first_run.text = first_run_text[:first_char_offset] + new_text

        # 2. Clear text from intermediate runs
        for run_idx in range(first_run_idx + 1, last_run_idx):
            runs[run_idx].text = ""

        # 3. Update the last run: keep text after the match
        last_run = runs[last_run_idx]
        last_run_text = last_run.text
        last_run.text = last_run_text[last_char_offset + 1 :]


def replace_in_paragraphs(
    paragraphs: Sequence[Paragraph],
    old_text: str,
    new_text: str,
    case_sensitive: bool = True,
    whole_word: bool = False,
) -> int:
    """Replace all occurrences of `old_text` with `new_text` across all paragraphs.

    Preserves the formatting of the first character's run for each replacement.
    Returns the total number of replacements made.
    """
    total = 0
    for paragraph in paragraphs:
        total += _replace_in_paragraph(
            paragraph, old_text, new_text, case_sensitive, whole_word
        )
    return total
