"""Collection providing access to endnotes in this document."""

from __future__ import annotations

from typing import TYPE_CHECKING
from collections.abc import Iterator

from docx.blkcntnr import BlockItemContainer
from docx.enum.text import WD_ENDNOTE_POSITION, WD_FOOTNOTE_RESTART, WD_NUMBER_FORMAT

if TYPE_CHECKING:
    from docx.oxml.endnotes import CT_EdnDocProps, CT_Endnote, CT_Endnotes
    from docx.parts.endnotes import EndnotesPart
    from docx.styles.style import ParagraphStyle
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run


class Endnotes:
    """Collection containing the endnotes in this document.

    .. versionadded:: 2026.05.0
    """

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

        .. versionadded:: 2026.05.0
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

    .. versionadded:: 2026.05.0
    """

    def __init__(self, endnote_elm: CT_Endnote, endnotes_part: EndnotesPart):
        super().__init__(endnote_elm, endnotes_part)
        self._endnote_elm = endnote_elm

    def clear(self) -> Endnote:
        """Remove all content from this endnote, leaving a single empty paragraph.

        The empty paragraph has the "EndnoteText" style. Returns this same endnote
        object for fluent use.

        .. versionadded:: 2026.05.0
        """
        self._endnote_elm.clear_content()
        return self

    def delete(self) -> None:
        """Remove this endnote from the document.

        Removes the `w:endnoteReference` element from the document body that references
        this endnote, along with the run containing it (if the run becomes empty). Also
        removes the `w:endnote` element from the endnotes part.

        After calling this method, this |Endnote| object is "defunct" and should not be
        used further.

        .. versionadded:: 2026.05.0
        """
        endnote_id = self.endnote_id
        # -- remove endnoteReference(s) from the document body --
        document_elm = self.part._document_part.element  # pyright: ignore[reportPrivateUsage]
        refs = document_elm.xpath(
            f'.//w:endnoteReference[@w:id="{endnote_id}"]',
        )
        for ref in refs:
            r = ref.getparent()
            if r is None:
                continue
            r.remove(ref)
            # -- remove the run if it's now empty (only rPr or nothing left) --
            if len(r.xpath("./*[not(self::w:rPr)]")) == 0:
                r_parent = r.getparent()
                if r_parent is not None:
                    r_parent.remove(r)
        # -- remove the endnote element from the endnotes part --
        endnotes_elm = self._endnote_elm.getparent()
        if endnotes_elm is not None:
            endnotes_elm.remove(self._endnote_elm)

    def add_paragraph(self, text: str = "", style: str | ParagraphStyle | None = None) -> Paragraph:
        """Return paragraph newly added to the end of the content in this container.

        The paragraph has `text` in a single run if present, and is given paragraph style `style`.
        When `style` is |None| or omitted, the "EndnoteText" paragraph style is applied, which is
        the default style for endnotes.

        .. versionadded:: 2026.05.0
        """
        paragraph = super().add_paragraph(text, style)

        if style is None:
            paragraph._p.style = "EndnoteText"  # pyright: ignore[reportPrivateUsage]

        return paragraph

    @property
    def endnote_id(self) -> int:
        """The unique identifier of this endnote.

        .. versionadded:: 2026.05.0
        """
        return self._endnote_elm.id

    @property
    def text(self) -> str:
        """The text content of this endnote as a string.

        Only content in paragraphs is included and all emphasis and styling is stripped.

        Paragraph boundaries are indicated with a newline (`"\\n"`).

        .. versionadded:: 2026.05.0
        """
        return "\n".join(p.text for p in self.paragraphs)


class EndnoteProperties:
    """Proxy for a ``<w:endnotePr>`` element providing endnote configuration.

    A `w:endnotePr` element can appear either at document level (as a child of
    `w:settings`) or at section level (as a child of `w:sectPr`). In either case it
    specifies the number format, position, starting number, and restart behaviour for
    endnote numbering.

    All properties return `None` when the corresponding child element is absent.
    Assigning `None` to a property removes the corresponding child element.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, endnotePr: "CT_EdnDocProps"):
        self._endnotePr = endnotePr

    @property
    def element(self) -> "CT_EdnDocProps":
        """The underlying ``<w:endnotePr>`` XML element.

        .. versionadded:: 2026.05.0
        """
        return self._endnotePr

    @property
    def number_format(self) -> WD_NUMBER_FORMAT | None:
        """The :ref:`WdNumberFormat` member corresponding to ``w:numFmt/@w:val``.

        Read/write. Returns |None| when no ``w:numFmt`` child is present.

        .. versionadded:: 2026.05.0
        """
        numFmt = self._endnotePr.numFmt
        if numFmt is None:
            return None
        return numFmt.val

    @number_format.setter
    def number_format(self, value: WD_NUMBER_FORMAT | None):
        if value is None:
            self._endnotePr._remove_numFmt()  # pyright: ignore[reportPrivateUsage]
            return
        numFmt = self._endnotePr.get_or_add_numFmt()
        numFmt.val = value

    @property
    def start_number(self) -> int | None:
        """The initial endnote number from ``w:numStart/@w:val`` as an int.

        Read/write. Returns |None| when no ``w:numStart`` child is present.

        .. versionadded:: 2026.05.0
        """
        numStart = self._endnotePr.numStart
        if numStart is None:
            return None
        return numStart.val

    @start_number.setter
    def start_number(self, value: int | None):
        if value is None:
            self._endnotePr._remove_numStart()  # pyright: ignore[reportPrivateUsage]
            return
        numStart = self._endnotePr.get_or_add_numStart()
        numStart.val = value

    @property
    def restart_rule(self) -> WD_FOOTNOTE_RESTART | None:
        """The :ref:`WdFootnoteRestart` member indicating when numbering restarts.

        Read/write. Corresponds to ``w:numRestart/@w:val``. Returns |None| when no
        ``w:numRestart`` child is present. Note that only ``CONTINUOUS`` and
        ``EACH_SECTION`` are meaningful for endnote numbering.

        .. versionadded:: 2026.05.0
        """
        numRestart = self._endnotePr.numRestart
        if numRestart is None:
            return None
        return numRestart.val

    @restart_rule.setter
    def restart_rule(self, value: WD_FOOTNOTE_RESTART | None):
        if value is None:
            self._endnotePr._remove_numRestart()  # pyright: ignore[reportPrivateUsage]
            return
        numRestart = self._endnotePr.get_or_add_numRestart()
        numRestart.val = value

    @property
    def position(self) -> WD_ENDNOTE_POSITION | None:
        """The :ref:`WdEndnotePosition` member indicating where endnotes appear.

        Read/write. Corresponds to ``w:pos/@w:val``. Returns |None| when no ``w:pos``
        child is present.

        .. versionadded:: 2026.05.0
        """
        pos = self._endnotePr.pos
        if pos is None or pos.val is None:
            return None
        return WD_ENDNOTE_POSITION.from_xml(pos.val)

    @position.setter
    def position(self, value: WD_ENDNOTE_POSITION | None):
        if value is None:
            self._endnotePr._remove_pos()  # pyright: ignore[reportPrivateUsage]
            return
        pos = self._endnotePr.get_or_add_pos()
        pos.val = WD_ENDNOTE_POSITION.to_xml(value)
