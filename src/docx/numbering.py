"""Numbering-related proxy objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, List

from docx.oxml.numbering import CT_AbstractNum, CT_Lvl, CT_Num, CT_Numbering
from docx.shared import ElementProxy

if TYPE_CHECKING:
    from docx.oxml.ns import qn


class Numbering(ElementProxy):
    """Proxy for the ``<w:numbering>`` element, providing access to numbering
    definitions in a document."""

    def __init__(self, numbering_elm: CT_Numbering, part: object):
        super().__init__(numbering_elm)
        self._numbering = numbering_elm
        self._part = part

    @property
    def definitions(self) -> List[NumberingDefinition]:
        """All numbering definitions (``<w:num>`` elements) in this numbering part."""
        return [
            NumberingDefinition(num, self._numbering)
            for num in self._numbering.num_lst
        ]

    def add_numbering_definition(
        self, levels: List[dict] | None = None
    ) -> NumberingDefinition:
        """Create a new numbering definition with the specified level formats.

        `levels` is a list of dicts, each specifying a level's format. Each dict can
        contain:

        - ``"level"`` (int): the level index 0-8 (default: position in list)
        - ``"format"`` (str): number format, e.g. "decimal", "lowerAlpha",
          "upperRoman", "bullet" (default: "decimal")
        - ``"text"`` (str): level text pattern, e.g. "%1.", "%1.%2"
          (default: "%{level+1}.")
        - ``"start"`` (int): starting number (default: 1)

        Returns a |NumberingDefinition| that can be applied to paragraphs.
        """
        numbering = self._numbering
        abstract_num_id = numbering._next_abstractNumId
        abstract_num = CT_AbstractNum.new(abstract_num_id)

        if levels is None:
            levels = [{"format": "decimal", "text": "%1.", "start": 1}]

        for idx, level_spec in enumerate(levels):
            ilvl = level_spec.get("level", idx)
            fmt = level_spec.get("format", "decimal")
            text = level_spec.get("text", "%%%d." % (ilvl + 1))
            start = level_spec.get("start", 1)

            lvl = abstract_num.add_lvl(ilvl)
            lvl.start_val = start
            lvl.numFmt_val = fmt
            lvl.lvlText_val = text

        numbering.add_abstractNum(abstract_num)
        num = numbering.add_num(abstract_num_id)
        return NumberingDefinition(num, numbering)


class NumberingDefinition(ElementProxy):
    """Proxy for a ``<w:num>`` element representing a concrete numbering definition."""

    def __init__(self, num: CT_Num, numbering: CT_Numbering):
        super().__init__(num)
        self._num = num
        self._numbering = numbering

    @property
    def num_id(self) -> int:
        """The ``numId`` of this numbering definition."""
        return self._num.numId

    @property
    def abstract_num_id(self) -> int:
        """The ``abstractNumId`` referenced by this definition."""
        return self._num.abstractNumId_val

    @property
    def levels(self) -> List[LevelFormat]:
        """The level format objects for this numbering definition."""
        abstract_num = self._numbering.abstractNum_having_abstractNumId(
            self.abstract_num_id
        )
        return [LevelFormat(lvl) for lvl in abstract_num.lvl_lst]

    def restart(self) -> NumberingDefinition:
        """Create a new numbering definition that restarts numbering at 1.

        Returns the new |NumberingDefinition| with a level override that sets
        ``<w:startOverride>`` to 1.
        """
        numbering = self._numbering
        new_num = numbering.add_num(self.abstract_num_id)
        lvl_override = new_num.add_lvlOverride(ilvl=0)
        lvl_override.add_startOverride(val=1)
        return NumberingDefinition(new_num, numbering)


class LevelFormat(ElementProxy):
    """Proxy for a ``<w:lvl>`` element describing the format at one list level."""

    def __init__(self, lvl: CT_Lvl):
        super().__init__(lvl)
        self._lvl = lvl

    @property
    def level(self) -> int:
        """The level index (0-8) of this format."""
        return self._lvl.ilvl

    @property
    def number_format(self) -> str | None:
        """The number format string (e.g. "decimal", "bullet")."""
        return self._lvl.numFmt_val

    @property
    def text_pattern(self) -> str | None:
        """The level text pattern (e.g. "%1.")."""
        return self._lvl.lvlText_val

    @property
    def start(self) -> int:
        """The starting number for this level."""
        return self._lvl.start_val
