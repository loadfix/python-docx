"""Collection providing access to footnotes in this document."""

from __future__ import annotations

from typing import TYPE_CHECKING
from collections.abc import Iterator

from docx.blkcntnr import BlockItemContainer
from docx.enum.text import WD_FOOTNOTE_POSITION, WD_FOOTNOTE_RESTART, WD_NUMBER_FORMAT

if TYPE_CHECKING:
    from docx.oxml.footnotes import CT_Footnote, CT_Footnotes, CT_FtnDocProps
    from docx.parts.footnotes import FootnotesPart
    from docx.styles.style import ParagraphStyle
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run


class Footnotes:
    """Collection containing the footnotes in this document.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, footnotes_elm: CT_Footnotes, footnotes_part: FootnotesPart):
        self._footnotes_elm = footnotes_elm
        self._footnotes_part = footnotes_part

    def __iter__(self) -> Iterator[Footnote]:
        return (
            Footnote(footnote_elm, self._footnotes_part)
            for footnote_elm in self._footnotes_elm.footnote_lst
            if footnote_elm.type is None
        )

    def __len__(self) -> int:
        return sum(1 for fn in self._footnotes_elm.footnote_lst if fn.type is None)

    def add(self, run: Run, text: str = "") -> Footnote:
        """Add a new footnote referenced from `run` and return it.

        A `w:footnoteReference` element is inserted into `run`, styled with the
        "FootnoteReference" character style. The new footnote contains a single paragraph
        with the "FootnoteText" style. If `text` is provided, it is added as a run in that
        paragraph following the footnote reference mark.

        .. versionadded:: 2026.05.0
        """
        footnote_elm = self._footnotes_elm.add_footnote()
        footnote = Footnote(footnote_elm, self._footnotes_part)

        # -- insert footnoteReference into the specified run in the document body --
        run._r.insert_footnote_reference(footnote_elm.id)  # pyright: ignore[reportPrivateUsage]

        # -- add text to the first paragraph if provided --
        if text:
            first_para = footnote.paragraphs[0]
            first_para.add_run(text)

        return footnote


class Footnote(BlockItemContainer):
    """Proxy for a single footnote in the document.

    A footnote is a block-item container, similar to a table cell, so it can contain both
    paragraphs and tables.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, footnote_elm: CT_Footnote, footnotes_part: FootnotesPart):
        super().__init__(footnote_elm, footnotes_part)
        self._footnote_elm = footnote_elm

    def clear(self) -> Footnote:
        """Remove all content from this footnote, leaving a single empty paragraph.

        The empty paragraph has the "FootnoteText" style. Returns this same footnote
        object for fluent use.

        .. versionadded:: 2026.05.0
        """
        self._footnote_elm.clear_content()
        return self

    def delete(self) -> None:
        """Remove this footnote from the document.

        Removes the `w:footnoteReference` element from the document body that references
        this footnote, along with the run containing it (if the run becomes empty). Also
        removes the `w:footnote` element from the footnotes part.

        After calling this method, this |Footnote| object is "defunct" and should not be
        used further.

        .. versionadded:: 2026.05.0
        """
        footnote_id = self.footnote_id
        # -- remove footnoteReference(s) from the document body --
        document_elm = self.part._document_part.element  # pyright: ignore[reportPrivateUsage]
        refs = document_elm.xpath(
            f'.//w:footnoteReference[@w:id="{footnote_id}"]',
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
        # -- remove the footnote element from the footnotes part --
        footnotes_elm = self._footnote_elm.getparent()
        if footnotes_elm is not None:
            footnotes_elm.remove(self._footnote_elm)

    def add_paragraph(self, text: str = "", style: str | ParagraphStyle | None = None) -> Paragraph:
        """Return paragraph newly added to the end of the content in this container.

        The paragraph has `text` in a single run if present, and is given paragraph style `style`.
        When `style` is |None| or omitted, the "FootnoteText" paragraph style is applied, which is
        the default style for footnotes.

        .. versionadded:: 2026.05.0
        """
        paragraph = super().add_paragraph(text, style)

        if style is None:
            paragraph._p.style = "FootnoteText"  # pyright: ignore[reportPrivateUsage]

        return paragraph

    @property
    def footnote_id(self) -> int:
        """The unique identifier of this footnote.

        .. versionadded:: 2026.05.0
        """
        return self._footnote_elm.id

    @property
    def text(self) -> str:
        """The text content of this footnote as a string.

        Only content in paragraphs is included and all emphasis and styling is stripped.

        Paragraph boundaries are indicated with a newline (`"\\n"`).

        .. versionadded:: 2026.05.0
        """
        return "\n".join(p.text for p in self.paragraphs)


class FootnoteProperties:
    """Proxy for a ``<w:footnotePr>`` element providing footnote configuration.

    A `w:footnotePr` element can appear either at document level (as a child of
    `w:settings`) or at section level (as a child of `w:sectPr`). In either case it
    specifies the number format, position, starting number, and restart behaviour for
    footnote numbering.

    All properties return `None` when the corresponding child element is absent.
    Assigning `None` to a property removes the corresponding child element.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, footnotePr: "CT_FtnDocProps"):
        self._footnotePr = footnotePr

    @property
    def element(self) -> "CT_FtnDocProps":
        """The underlying ``<w:footnotePr>`` XML element.

        .. versionadded:: 2026.05.0
        """
        return self._footnotePr

    @property
    def number_format(self) -> WD_NUMBER_FORMAT | None:
        """The :ref:`WdNumberFormat` member corresponding to ``w:numFmt/@w:val``.

        Read/write. Returns |None| when no ``w:numFmt`` child is present.

        .. versionadded:: 2026.05.0
        """
        numFmt = self._footnotePr.numFmt
        if numFmt is None:
            return None
        return numFmt.val

    @number_format.setter
    def number_format(self, value: WD_NUMBER_FORMAT | None):
        if value is None:
            self._footnotePr._remove_numFmt()  # pyright: ignore[reportPrivateUsage]
            return
        numFmt = self._footnotePr.get_or_add_numFmt()
        numFmt.val = value

    @property
    def start_number(self) -> int | None:
        """The initial footnote number from ``w:numStart/@w:val`` as an int.

        Read/write. Returns |None| when no ``w:numStart`` child is present.

        .. versionadded:: 2026.05.0
        """
        numStart = self._footnotePr.numStart
        if numStart is None:
            return None
        return numStart.val

    @start_number.setter
    def start_number(self, value: int | None):
        if value is None:
            self._footnotePr._remove_numStart()  # pyright: ignore[reportPrivateUsage]
            return
        numStart = self._footnotePr.get_or_add_numStart()
        numStart.val = value

    @property
    def restart_rule(self) -> WD_FOOTNOTE_RESTART | None:
        """The :ref:`WdFootnoteRestart` member indicating when numbering restarts.

        Read/write. Corresponds to ``w:numRestart/@w:val``. Returns |None| when no
        ``w:numRestart`` child is present.

        .. versionadded:: 2026.05.0
        """
        numRestart = self._footnotePr.numRestart
        if numRestart is None:
            return None
        return numRestart.val

    @restart_rule.setter
    def restart_rule(self, value: WD_FOOTNOTE_RESTART | None):
        if value is None:
            self._footnotePr._remove_numRestart()  # pyright: ignore[reportPrivateUsage]
            return
        numRestart = self._footnotePr.get_or_add_numRestart()
        numRestart.val = value

    @property
    def position(self) -> WD_FOOTNOTE_POSITION | None:
        """The :ref:`WdFootnotePosition` member indicating footnote page position.

        Read/write. Corresponds to ``w:pos/@w:val``. Returns |None| when no ``w:pos``
        child is present.

        .. versionadded:: 2026.05.0
        """
        pos = self._footnotePr.pos
        if pos is None or pos.val is None:
            return None
        return WD_FOOTNOTE_POSITION.from_xml(pos.val)

    @position.setter
    def position(self, value: WD_FOOTNOTE_POSITION | None):
        if value is None:
            self._footnotePr._remove_pos()  # pyright: ignore[reportPrivateUsage]
            return
        pos = self._footnotePr.get_or_add_pos()
        pos.val = WD_FOOTNOTE_POSITION.to_xml(value)
