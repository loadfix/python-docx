"""Caption-building helpers.

A Word "caption" is a paragraph styled with the ``Caption`` style and built
around a ``SEQ`` field. The SEQ field auto-numbers captions in order by
*label* (e.g. ``Figure`` or ``Table``) so that Word can maintain the numbering
as captions are added or deleted.

The XML shape produced by this module looks like::

    <w:p>
      <w:pPr><w:pStyle w:val="Caption"/></w:pPr>
      <w:r><w:t xml:space="preserve">Figure </w:t></w:r>
      <w:fldSimple w:instr=' SEQ Figure \\* ARABIC '>
        <w:r><w:t>1</w:t></w:r>
      </w:fldSimple>
      <w:r><w:t xml:space="preserve">: </w:t></w:r>
      <w:r><w:t>A diagram of the system</w:t></w:r>
    </w:p>

The ``1`` inside the ``<w:fldSimple>`` is the cached field result; Word
recomputes it whenever the document is opened or fields are refreshed.

This module exposes a single low-level helper,
:func:`new_caption_paragraph`, which populates a freshly-created paragraph
with the caption structure. The public API is surfaced via
:meth:`docx.document.Document.add_caption`,
:meth:`docx.text.paragraph.Paragraph.add_caption_before`, and
:meth:`docx.text.paragraph.Paragraph.add_caption_after`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docx.text.paragraph import Paragraph


def new_caption_paragraph(
    paragraph: Paragraph,
    text: str,
    label: str = "Figure",
    style: str = "Caption",
) -> Paragraph:
    """Populate `paragraph` as a caption of the form ``"{label} N: {text}"``.

    `paragraph` must be an empty, freshly-created |Paragraph|. The paragraph's
    style is set to `style` (defaulting to ``"Caption"``), and the standard
    caption run sequence — literal label, ``SEQ`` field, literal separator,
    and caption text — is appended.

    The caller is responsible for positioning the paragraph in the document.
    The populated paragraph is returned.

    .. versionadded:: 2026.05.0
    """
    paragraph.style = style
    # -- literal label plus trailing space (e.g. "Figure ") --
    paragraph.add_run(f"{label} ")
    # -- SEQ field for the auto-number; "1" is a cached result Word updates --
    paragraph.add_simple_field(f" SEQ {label} \\* ARABIC ", "1")
    # -- literal ": " separator before the caption text --
    paragraph.add_run(": ")
    # -- caption text --
    if text:
        paragraph.add_run(text)
    return paragraph
