"""Readability metrics for python-docx documents.

Closes #58. Computes the six standard readability scores -- Flesch Reading
Ease, Flesch-Kincaid Grade, Gunning Fog, SMOG, Coleman-Liau, and Automated
Readability Index -- plus the underlying word / sentence / syllable /
complex-word counts, both for the whole body story and for each
heading-bounded section.

A "section" is everything from one ``Heading 1`` paragraph up to (but not
including) the next ``Heading 1``. Body content before the first
``Heading 1`` is gathered into a synthetic ``Introduction`` section.
``Title``-styled paragraphs and headings of level 2..9 do *not* split
sections -- they are folded into the surrounding ``Heading 1`` section.

Three-line usage example::

    from docx import Document
    metrics = Document("paper.docx").readability()
    print(metrics.overall.flesch_kincaid_grade)

Algorithm notes
---------------

* **Sentences** are split on the regex ``[.!?]+`` followed by whitespace
  or end-of-string; this is a heuristic but matches the textstat /
  py-readability-metrics conventions for Flesch family scores.
* **Words** are whitespace-delimited tokens stripped of leading and
  trailing punctuation. Numbers are counted as words. A token must
  contain at least one alphanumeric character to count.
* **Syllables** are counted with a vowel-group heuristic: each maximal
  contiguous run of vowels (``aeiouy``) is one syllable, with a trailing
  silent ``e`` discounted (except in monosyllables like ``the``). Every
  word has at least one syllable. The heuristic is good enough for the
  formulas; published scores are tolerant of small drift.
* **Complex words** (used by the Gunning Fog and SMOG formulas) are
  three-or-more-syllable words that are *not* proper nouns (capitalised
  in mid-sentence) and *not* familiar suffix-driven compounds (a token
  ending in ``-es``/``-ed``/``-ing`` is reduced to its stem before the
  syllable count is taken). This matches Gunning's original heuristic.

Formulas (all standard, taken from the Wikipedia reference page each
metric links to):

* Flesch Reading Ease   = 206.835 - 1.015 * (W/S) - 84.6 * (Syl/W)
* Flesch-Kincaid Grade  = 0.39   * (W/S) + 11.8   * (Syl/W) - 15.59
* Gunning Fog           = 0.4    * (W/S + 100 * complex/W)
* SMOG                  = 1.0430 * sqrt(complex * 30 / S) + 3.1291
* Coleman-Liau          = 0.0588 * L  - 0.296 * S100 - 15.8
                          (L = letters per 100 words, S100 = sentences
                          per 100 words)
* Automated Readability = 4.71 * (chars/W) + 0.5 * (W/S) - 21.43

All formulas return ``0.0`` when the document has no words or no
sentences, rather than raising ``ZeroDivisionError``.

.. versionadded:: 2026.05.12
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, List, Optional, Tuple

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph


# -- "Heading 1" splits sections; "Title" / "Heading 2..9" do not. --
_H1_RE = re.compile(r"^heading\s+1$", re.IGNORECASE)

# -- Sentence splitter: end-of-sentence punctuation followed by whitespace
# -- or end-of-string. Doesn't try to handle abbreviations -- the formulas
# -- are tolerant of a few extra splits. --
_SENTENCE_SPLIT_RE = re.compile(r"[.!?]+(?:\s+|$)")

# -- Word tokenizer: whitespace splits, then we strip surrounding
# -- punctuation. Tokens that contain no alphanumerics (e.g. "--", "...")
# -- don't count as words. --
_WORD_PUNCT_STRIP_RE = re.compile(r"^[^\w]+|[^\w]+$", re.UNICODE)

_VOWELS = "aeiouy"


def _strip_word(token: str) -> str:
    """Strip leading/trailing punctuation from `token`.

    Returns the bare word (or the empty string if `token` was pure
    punctuation).
    """
    return _WORD_PUNCT_STRIP_RE.sub("", token)


def _is_word(token: str) -> bool:
    """Return True if `token` is a real word (contains an alphanumeric)."""
    return any(ch.isalnum() for ch in token)


def count_syllables(word: str) -> int:
    """Return an estimate of the number of syllables in `word`.

    Vowel-group heuristic: count contiguous runs of vowels, drop one for
    a silent trailing ``e`` (except when that would zero the count, as
    in ``the`` or ``be``). Every real word has at least one syllable.

    The heuristic is intentionally simple -- the published Flesch /
    Gunning / SMOG scores absorb small per-word drift because the
    formulas average over hundreds of words.
    """
    word = word.lower()
    # -- strip non-alpha (digits, hyphens) -- numbers count as one syllable. --
    letters = "".join(ch for ch in word if ch.isalpha())
    if not letters:
        # -- pure-digit token -- count as one syllable (a digit is read aloud).
        return 1 if word else 0
    count = 0
    prev_was_vowel = False
    for ch in letters:
        is_vowel = ch in _VOWELS
        if is_vowel and not prev_was_vowel:
            count += 1
        prev_was_vowel = is_vowel
    # -- silent trailing e: "house" -> 1, not 2; but "be" stays at 1. --
    if letters.endswith("e") and count > 1:
        count -= 1
    # -- guard: every real word is at least one syllable. --
    return max(count, 1)


def _stem_for_complex(word: str) -> str:
    """Return `word` with a Gunning-style suffix stripped, lowercased.

    Strips ``-es``, ``-ed``, ``-ing`` (the suffixes Gunning called out
    as "easy" so they don't push a word into the complex bucket). This
    is *only* used for the complex-word check; the syllable count for
    the Flesch family uses the unstripped word.
    """
    w = word.lower()
    for suffix in ("ing", "ed", "es"):
        if w.endswith(suffix) and len(w) > len(suffix) + 1:
            return w[: -len(suffix)]
    return w


def _is_complex(word: str, mid_sentence: bool) -> bool:
    """Return True if `word` is a "complex word" per Gunning Fog.

    A complex word is one with three or more syllables that is not a
    proper noun (capitalised in mid-sentence) and whose stem (after
    stripping ``-es``/``-ed``/``-ing``) still has three or more
    syllables.
    """
    bare = word.strip()
    if not bare:
        return False
    # -- proper nouns: capitalised tokens in mid-sentence position. --
    if mid_sentence and bare[0].isupper():
        return False
    stem = _stem_for_complex(bare)
    return count_syllables(stem) >= 3


def _split_sentences(text: str) -> List[str]:
    """Return non-empty sentence strings split out of `text`."""
    if not text:
        return []
    parts = _SENTENCE_SPLIT_RE.split(text)
    return [p for p in (s.strip() for s in parts) if p]


def _tokenize_words(text: str) -> List[str]:
    """Return the list of bare-word tokens from `text` (in document order)."""
    raw = text.split()
    out: List[str] = []
    for tok in raw:
        bare = _strip_word(tok)
        if _is_word(bare):
            out.append(bare)
    return out


@dataclass
class TextCounts:
    """Aggregated counts for one body of text.

    Used internally by :func:`compute_metrics`; not exposed publicly.
    """

    words: int = 0
    sentences: int = 0
    syllables: int = 0
    characters: int = 0  # -- letters and digits only, used by Coleman-Liau --
    complex_words: int = 0


def _accumulate(counts: TextCounts, text: str) -> None:
    """Fold `text`'s tokens into `counts` in place."""
    sentences = _split_sentences(text)
    counts.sentences += len(sentences)
    for sentence in sentences:
        tokens = _tokenize_words(sentence)
        for pos, token in enumerate(tokens):
            counts.words += 1
            counts.syllables += count_syllables(token)
            counts.characters += sum(1 for ch in token if ch.isalnum())
            if _is_complex(token, mid_sentence=pos > 0):
                counts.complex_words += 1


def _safe_div(numerator: float, denominator: float) -> float:
    """Return ``numerator / denominator`` or ``0.0`` when the divisor is zero."""
    if denominator == 0:
        return 0.0
    return numerator / denominator


@dataclass
class ReadabilityScores:
    """All readability metrics + counts for one body of text.

    Returned as both ``ReadabilityReport.overall`` and as the
    ``scores`` of each ``SectionScores``.

    Attributes:

    * ``flesch_reading_ease`` -- Flesch Reading Ease (0-100 typical;
      higher = easier).
    * ``flesch_kincaid_grade`` -- US-grade-level estimate (Flesch-Kincaid).
    * ``gunning_fog`` -- US-grade-level estimate (Gunning Fog).
    * ``smog`` -- US-grade-level estimate (SMOG).
    * ``coleman_liau`` -- US-grade-level estimate (Coleman-Liau).
    * ``automated_readability`` -- US-grade-level estimate (ARI).
    * ``word_count`` -- number of word tokens (post-punctuation strip).
    * ``sentence_count`` -- number of sentences (heuristic split).
    * ``syllable_count`` -- total syllables across all words.
    * ``character_count`` -- alphanumeric characters across all words.
    * ``complex_word_count`` -- words with >=3 syllables, ignoring
      proper nouns and ``-ing``/``-ed``/``-es`` suffix forms.
    * ``avg_words_per_sentence`` -- words / sentences (0 when no sentences).
    * ``avg_syllables_per_word`` -- syllables / words (0 when no words).

    .. versionadded:: 2026.05.12
    """

    flesch_reading_ease: float = 0.0
    flesch_kincaid_grade: float = 0.0
    gunning_fog: float = 0.0
    smog: float = 0.0
    coleman_liau: float = 0.0
    automated_readability: float = 0.0
    word_count: int = 0
    sentence_count: int = 0
    syllable_count: int = 0
    character_count: int = 0
    complex_word_count: int = 0
    avg_words_per_sentence: float = 0.0
    avg_syllables_per_word: float = 0.0


@dataclass
class SectionScores:
    """Readability scores for one heading-bounded section.

    ``title`` is the section's heading text (or ``"Introduction"`` for
    the synthetic pre-first-heading section). ``paragraph_index`` is
    the position of the section's heading paragraph in
    :attr:`Document.paragraphs` -- ``-1`` for the synthetic
    introduction section (no heading paragraph backs it).

    Each metric and count is mirrored on the section for convenient
    access (``section.flesch_kincaid_grade`` works as well as
    ``section.scores.flesch_kincaid_grade``).

    .. versionadded:: 2026.05.12
    """

    title: str
    paragraph_index: int
    scores: ReadabilityScores = field(default_factory=ReadabilityScores)

    def __getattr__(self, name: str):
        # -- Convenience pass-through so callers can write
        # -- ``section.flesch_kincaid_grade`` per the issue example.
        # -- ``__getattr__`` is only invoked for attributes not found
        # -- on the dataclass itself, so ``title`` / ``paragraph_index``
        # -- / ``scores`` resolve normally. --
        try:
            scores = object.__getattribute__(self, "scores")
        except AttributeError:
            raise AttributeError(name)
        if hasattr(scores, name):
            return getattr(scores, name)
        raise AttributeError(name)


@dataclass
class ReadabilityReport:
    """Whole-document readability snapshot.

    Returned by :meth:`docx.document.Document.readability`.

    Attributes:

    * ``overall`` -- :class:`ReadabilityScores` for the entire body
      story (every paragraph treated as one block of text).
    * ``sections`` -- one :class:`SectionScores` per heading-bounded
      section. The first entry is the synthetic ``Introduction``
      section iff there is body text before the first ``Heading 1``;
      otherwise sections start at the first ``Heading 1``. Sections
      with zero words are still included so downstream callers can
      detect empty headings.

    .. versionadded:: 2026.05.12
    """

    overall: ReadabilityScores = field(default_factory=ReadabilityScores)
    sections: List[SectionScores] = field(default_factory=list)

    def to_dict(self) -> "dict[str, object]":
        """Return a JSON-serialisable snapshot of this report.

        .. versionadded:: 2026.05.12
        """
        def scores_to_dict(s: ReadabilityScores) -> "dict[str, object]":
            return {
                "flesch_reading_ease": s.flesch_reading_ease,
                "flesch_kincaid_grade": s.flesch_kincaid_grade,
                "gunning_fog": s.gunning_fog,
                "smog": s.smog,
                "coleman_liau": s.coleman_liau,
                "automated_readability": s.automated_readability,
                "word_count": s.word_count,
                "sentence_count": s.sentence_count,
                "syllable_count": s.syllable_count,
                "character_count": s.character_count,
                "complex_word_count": s.complex_word_count,
                "avg_words_per_sentence": s.avg_words_per_sentence,
                "avg_syllables_per_word": s.avg_syllables_per_word,
            }

        return {
            "overall": scores_to_dict(self.overall),
            "sections": [
                {
                    "title": sec.title,
                    "paragraph_index": sec.paragraph_index,
                    **scores_to_dict(sec.scores),
                }
                for sec in self.sections
            ],
        }


def compute_metrics(text: str) -> ReadabilityScores:
    """Return a |ReadabilityScores| for `text`.

    `text` is treated as one continuous block; sentence splitting and
    word tokenisation happen internally.

    All six standard formulas are evaluated. When `text` has no words
    or no sentences, every metric collapses to ``0.0`` rather than
    raising.

    .. versionadded:: 2026.05.12
    """
    counts = TextCounts()
    _accumulate(counts, text)
    return _scores_from_counts(counts)


def _scores_from_counts(c: TextCounts) -> ReadabilityScores:
    """Materialise a |ReadabilityScores| from raw counts."""
    w, s = c.words, c.sentences
    syl, ch, cw = c.syllables, c.characters, c.complex_words

    asl = _safe_div(w, s)  # -- average sentence length (words / sentence) --
    asw = _safe_div(syl, w)  # -- average syllables per word --

    if w == 0 or s == 0:
        return ReadabilityScores(
            word_count=w,
            sentence_count=s,
            syllable_count=syl,
            character_count=ch,
            complex_word_count=cw,
            avg_words_per_sentence=asl,
            avg_syllables_per_word=asw,
        )

    flesch_re = 206.835 - 1.015 * asl - 84.6 * asw
    fk_grade = 0.39 * asl + 11.8 * asw - 15.59
    fog = 0.4 * (asl + 100.0 * _safe_div(cw, w))
    smog = 1.0430 * math.sqrt(_safe_div(cw * 30.0, s)) + 3.1291
    # -- Coleman-Liau: L = letters per 100 words, S = sentences per 100 words --
    L_per_100 = _safe_div(ch * 100.0, w)
    S_per_100 = _safe_div(s * 100.0, w)
    cli = 0.0588 * L_per_100 - 0.296 * S_per_100 - 15.8
    ari = 4.71 * _safe_div(ch, w) + 0.5 * asl - 21.43

    return ReadabilityScores(
        flesch_reading_ease=flesch_re,
        flesch_kincaid_grade=fk_grade,
        gunning_fog=fog,
        smog=smog,
        coleman_liau=cli,
        automated_readability=ari,
        word_count=w,
        sentence_count=s,
        syllable_count=syl,
        character_count=ch,
        complex_word_count=cw,
        avg_words_per_sentence=asl,
        avg_syllables_per_word=asw,
    )


def _is_h1(paragraph: "Paragraph") -> bool:
    """Return True if `paragraph` is styled ``Heading 1``."""
    style = paragraph.style
    if style is None:
        return False
    name = getattr(style, "name", None)
    if not isinstance(name, str):
        return False
    return _H1_RE.match(name.strip()) is not None


def _section_partition(
    paragraphs: "List[Paragraph]",
) -> List[Tuple[str, int, List[str]]]:
    """Partition body paragraphs into sections by ``Heading 1``.

    Returns a list of ``(title, paragraph_index, [paragraph_text, ...])``
    tuples. The first tuple is the synthetic ``Introduction`` section
    iff there is non-empty text before the first ``Heading 1``;
    otherwise sections start at the first heading.
    """
    sections: List[Tuple[str, int, List[str]]] = []
    intro_texts: List[str] = []
    current_title: Optional[str] = None
    current_index: int = -1
    current_texts: List[str] = []

    for idx, paragraph in enumerate(paragraphs):
        text = paragraph.text or ""
        if _is_h1(paragraph):
            # -- close out previous section / introduction block --
            if current_title is None:
                if any(t.strip() for t in intro_texts):
                    sections.append(("Introduction", -1, intro_texts))
            else:
                sections.append((current_title, current_index, current_texts))
            # -- start the new section. The heading's own text is part
            # -- of the section so it counts toward word_count. --
            current_title = text.strip() or "Untitled"
            current_index = idx
            current_texts = [text]
        else:
            if current_title is None:
                intro_texts.append(text)
            else:
                current_texts.append(text)

    # -- flush the trailing section / introduction --
    if current_title is None:
        if any(t.strip() for t in intro_texts):
            sections.append(("Introduction", -1, intro_texts))
    else:
        sections.append((current_title, current_index, current_texts))

    return sections


def build_report(document: "Document") -> ReadabilityReport:
    """Return a :class:`ReadabilityReport` for `document`.

    Walks :attr:`Document.paragraphs` once, partitioning by
    ``Heading 1`` boundaries. The whole-document scores are computed
    from the same paragraph stream so the overall counts are exactly
    the sum of the per-section counts (modulo sentence-split drift at
    section boundaries).

    .. versionadded:: 2026.05.12
    """
    paragraphs = list(document.paragraphs)

    # -- Overall: join every paragraph with newlines so sentence-splitter
    # -- treats inter-paragraph breaks as sentence boundaries when the
    # -- preceding paragraph ends without punctuation. --
    overall_text = "\n".join(p.text or "" for p in paragraphs)
    overall = compute_metrics(overall_text)

    sections: List[SectionScores] = []
    for title, idx, texts in _section_partition(paragraphs):
        section_text = "\n".join(texts)
        scores = compute_metrics(section_text)
        sections.append(
            SectionScores(title=title, paragraph_index=idx, scores=scores)
        )

    return ReadabilityReport(overall=overall, sections=sections)
