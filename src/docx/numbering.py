"""Numbering-related proxy types for document numbering/list control."""

from __future__ import annotations

from typing import TYPE_CHECKING, List

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement

if TYPE_CHECKING:
    from docx.oxml.numbering import CT_AbstractNum, CT_Lvl, CT_Num, CT_Numbering
    from docx.parts.numbering import NumberingPart


class Numbering:
    """Provides access to numbering definitions in the document.

    Accessible via ``document.numbering``.
    """

    def __init__(self, numbering_elm: CT_Numbering, numbering_part: NumberingPart):
        self._numbering_elm = numbering_elm
        self._numbering_part = numbering_part

    @property
    def definitions(self) -> List[NumberingDefinition]:
        """All numbering definitions (``<w:num>`` elements) in this document."""
        return [
            NumberingDefinition(num, self._numbering_elm)
            for num in self._numbering_elm.num_lst
        ]

    def add_numbering_definition(
        self, levels: List[dict] | None = None
    ) -> NumberingDefinition:
        """Create a custom multi-level numbering definition and return it.

        `levels` is an optional list of dicts, each specifying:
            - ``number_format``: str — e.g. "decimal", "lowerLetter", "upperRoman",
              "bullet" (default "decimal")
            - ``text``: str — level text pattern, e.g. "%1.", "%1.%2" (auto-generated
              if omitted)
            - ``start``: int — starting number (default 1)
            - ``indent``: int — left indent in twips (auto-calculated if omitted)
            - ``font``: str — font name for this level (optional, mainly for bullets)

        If `levels` is None, a single-level decimal list is created.
        """
        if levels is None:
            levels = [{"number_format": "decimal", "text": "%1.", "start": 1}]

        # -- create abstract numbering definition --
        abstractNum = self._numbering_elm.add_abstractNum()

        for i, level_spec in enumerate(levels):
            lvl = abstractNum.add_lvl(i)
            fmt = level_spec.get("number_format", "decimal")
            lvl.numFmt_val = fmt

            text = level_spec.get("text")
            if text is None:
                text = "%" + str(i + 1) + "."
            lvl.lvlText_val = text

            start = level_spec.get("start", 1)
            lvl.start_val = start

            indent = level_spec.get("indent")
            if indent is not None:
                pPr = lvl.get_or_add_pPr()
                ind = OxmlElement("w:ind")
                ind.attrib[qn("w:left")] = str(indent)
                ind.attrib[qn("w:hanging")] = str(min(indent, 360))
                pPr.append(ind)

            font = level_spec.get("font")
            if font is not None:
                rPr = lvl.get_or_add_rPr()
                rFonts = OxmlElement("w:rFonts")
                rFonts.attrib[qn("w:ascii")] = font
                rFonts.attrib[qn("w:hAnsi")] = font
                rPr.append(rFonts)

        # -- create concrete num referencing the abstract num --
        num = self._numbering_elm.add_num(abstractNum.abstractNumId)

        return NumberingDefinition(num, self._numbering_elm)


class NumberingDefinition:
    """Proxy for a ``<w:num>`` element — a concrete numbering definition instance."""

    def __init__(self, num: CT_Num, numbering_elm: CT_Numbering):
        self._num = num
        self._numbering_elm = numbering_elm

    @property
    def num_id(self) -> int:
        """The ``numId`` attribute value identifying this numbering definition."""
        return self._num.numId

    @property
    def abstract_num_id(self) -> int:
        """The ``abstractNumId`` referenced by this numbering definition."""
        return self._num.abstractNumId_val

    @property
    def level_formats(self) -> List[LevelFormat]:
        """The level format objects for the abstract numbering definition backing this
        concrete definition."""
        abstractNumId = self.abstract_num_id
        for abstractNum in self._numbering_elm.abstractNum_lst:
            if abstractNum.abstractNumId == abstractNumId:
                return [LevelFormat(lvl) for lvl in abstractNum.lvl_lst]
        return []


class LevelFormat:
    """Proxy for a ``<w:lvl>`` element within an abstract numbering definition."""

    def __init__(self, lvl: CT_Lvl):
        self._lvl = lvl

    @property
    def level_index(self) -> int:
        """The ``ilvl`` attribute — the zero-based level index."""
        return self._lvl.ilvl

    @property
    def number_format(self) -> str | None:
        """The number format, e.g. "decimal", "lowerLetter", "bullet"."""
        return self._lvl.numFmt_val

    @property
    def text(self) -> str | None:
        """The level text pattern, e.g. "%1.", "%1.%2"."""
        return self._lvl.lvlText_val

    @property
    def start(self) -> int | None:
        """The starting number for this level."""
        return self._lvl.start_val
