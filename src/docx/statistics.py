"""Word / character / paragraph counting for python-docx documents.

Provides :class:`DocumentStatistics`, a small named-tuple-like object summarizing the
text content of a document's main story (the ``w:body``). Counts match Word's
"Word Count" statistics as closely as practical:

* Only body text is considered. Text in headers, footers, footnotes, endnotes, and
  comments is not included (those live in separate OOXML parts).
* A "paragraph" is a non-empty ``w:p`` element (empty paragraphs used purely for
  spacing are not counted, matching Word).
* A "word" is a whitespace-delimited run of non-whitespace characters, consistent
  with ``str.split()`` semantics.
* "Characters" counts every character in the collected text including spaces.
* "Characters (no spaces)" excludes all whitespace characters.

Paragraphs nested inside tables or structured-document tags (content controls) are
included because they are part of the body story.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, NamedTuple

if TYPE_CHECKING:
    from docx.oxml.document import CT_Body


class DocumentStatistics(NamedTuple):
    """Summary of text statistics for a document's body.

    Returned by :attr:`docx.document.Document.statistics`.

    .. versionadded:: 1.3.0.dev0
    """

    paragraphs: int
    """Count of non-empty body paragraphs."""

    words: int
    """Count of whitespace-delimited tokens across all body paragraph text."""

    characters: int
    """Count of characters in body text, including spaces."""

    characters_no_spaces: int
    """Count of characters in body text, excluding whitespace."""

    pages: int | None = None
    """Cached page count (from ``docProps/app.xml``), or |None| if not recorded.

    python-docx does not lay the document out so it cannot derive a true page
    count. This field surfaces the ``<Pages>`` value written by the last
    application that saved the document (Word, etc.); :obj:`None` is returned
    when no extended-properties part is present or the property is absent.

    .. versionadded:: 1.3.0.dev0
    """


def compute_statistics(body: CT_Body, pages: int | None = None) -> DocumentStatistics:
    """Return a |DocumentStatistics| for the given ``w:body`` element.

    Descends into tables and other block containers so all paragraphs in the body
    story contribute to the counts. When `pages` is given it populates
    :attr:`DocumentStatistics.pages`; otherwise that field is |None|.

    .. versionadded:: 1.3.0.dev0
    """
    # -- collect every paragraph in the body, including those nested inside tables
    # -- and block-level structured-document tags
    paragraph_texts = [p.text for p in body.xpath(".//w:p")]

    paragraph_count = sum(1 for text in paragraph_texts if text.strip())

    word_count = 0
    character_count = 0
    character_no_spaces_count = 0
    for text in paragraph_texts:
        word_count += len(text.split())
        character_count += len(text)
        character_no_spaces_count += sum(1 for ch in text if not ch.isspace())

    return DocumentStatistics(
        paragraphs=paragraph_count,
        words=word_count,
        characters=character_count,
        characters_no_spaces=character_no_spaces_count,
        pages=pages,
    )
