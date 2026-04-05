"""Collection providing access to endnotes in this document."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from docx.blkcntnr import BlockItemContainer

if TYPE_CHECKING:
    from docx.oxml.endnotes import CT_Endnote, CT_Endnotes
    from docx.parts.endnotes import EndnotesPart
    from docx.styles.style import ParagraphStyle
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run


class Endnotes:
    """Collection containing the endnotes in this document."""

    def __init__(self, endnotes_elm: CT_Endnotes, endnotes_part: EndnotesPart):
        self._endnotes_elm = endnotes_elm
        self._endnotes_part = endnotes_part

    def __iter__(self) -> Iterator[Endnote]:
        return (
            Endnote(endnote_elm, self._endnotes_part)
            for endnote_elm in self._endnotes_elm.endnote_lst
            if endnote_elm.type is None
        )

    def __len__(self) -> int:
        return sum(1 for en in self._endnotes_elm.endnote_lst if en.type is None)

    def add(self, run: Run, text: str = "") -> Endnote:
        """Add a new endnote referenced from `run` and return it.

        A `w:endnoteReference` element is inserted into `run`, styled with the
        "EndnoteReference" character style. The new endnote contains a single paragraph
        with the "EndnoteText" style. If `text` is provided, it is added as a run in that
        paragraph following the endnote reference mark.
        """
        endnote_elm = self._endnotes_elm.add_endnote()
        endnote = Endnote(endnote_elm, self._endnotes_part)

        # -- insert endnoteReference into the specified run in the document body --
        run._r.insert_endnote_reference(endnote_elm.id)  # pyright: ignore[reportPrivateUsage]

        # -- add text to the first paragraph if provided --
        if text:
            first_para = endnote.paragraphs[0]
            first_para.add_run(text)

        return endnote


class Endnote(BlockItemContainer):
    """Proxy for a single endnote in the document.

    An endnote is a block-item container, similar to a table cell, so it can contain both
    paragraphs and tables.
    """

    def __init__(self, endnote_elm: CT_Endnote, endnotes_part: EndnotesPart):
        super().__init__(endnote_elm, endnotes_part)
        self._endnote_elm = endnote_elm

    def add_paragraph(self, text: str = "", style: str | ParagraphStyle | None = None) -> Paragraph:
        """Return paragraph newly added to the end of the content in this container.

        The paragraph has `text` in a single run if present, and is given paragraph style `style`.
        When `style` is |None| or omitted, the "EndnoteText" paragraph style is applied, which is
        the default style for endnotes.
        """
        paragraph = super().add_paragraph(text, style)

        if style is None:
            paragraph._p.style = "EndnoteText"  # pyright: ignore[reportPrivateUsage]

        return paragraph

    @property
    def endnote_id(self) -> int:
        """The unique identifier of this endnote."""
        return self._endnote_elm.id
