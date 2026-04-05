"""List-format related proxy objects for paragraphs."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docx.numbering import NumberingDefinition
    from docx.oxml.text.paragraph import CT_P
    from docx.parts.document import DocumentPart


class ListFormat:
    """Provides access to the list formatting properties of a paragraph.

    Accessed via ``paragraph.list_format``.
    """

    def __init__(self, p: CT_P, part: DocumentPart):
        self._p = p
        self._part = part

    @property
    def level(self) -> int | None:
        """The list indentation level (0-8) for this paragraph.

        Returns None if this paragraph is not part of a list.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        numPr = pPr.numPr
        if numPr is None:
            return None
        return numPr.ilvl_val

    @level.setter
    def level(self, value: int | None) -> None:
        if value is None:
            pPr = self._p.pPr
            if pPr is not None and pPr.numPr is not None:
                pPr.numPr.ilvl_val = None
            return
        pPr = self._p.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        numPr.ilvl_val = value

    @property
    def num_id(self) -> int | None:
        """The numbering definition ID applied to this paragraph.

        Returns None if this paragraph is not part of a list.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        numPr = pPr.numPr
        if numPr is None:
            return None
        return numPr.numId_val

    @num_id.setter
    def num_id(self, value: int | None) -> None:
        if value is None:
            pPr = self._p.pPr
            if pPr is not None and pPr.numPr is not None:
                pPr.numPr.numId_val = None
            return
        pPr = self._p.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        numPr.numId_val = value

    def apply(self, num_id: int, level: int = 0) -> None:
        """Apply a numbering definition to this paragraph.

        `num_id` is the numId of the numbering definition. `level` is the
        indentation level (0-8, default 0).
        """
        pPr = self._p.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        numPr.numId_val = num_id
        numPr.ilvl_val = level
