# pyright: reportPrivateUsage=false

"""Proxy objects providing a convenient API over the numbering part.

Numbering (lists) in WordprocessingML are split into two parts:

- ``w:abstractNum`` elements describe the *visual* formatting of a list — its
  numbering format (decimal, bullet, Roman, etc.), its level text pattern,
  indentation, and font.
- ``w:num`` elements are *instances* that point at an abstract definition and
  can optionally override its starting number on a per-level basis.

A paragraph joins a list by pointing at a ``w:num`` (its ``numId``) and
declaring the level at which it should appear (its ``ilvl``).

This module exposes three proxies:

- :class:`Numbering` — the top-level collection (available as
  ``document.numbering``). Use :meth:`Numbering.add_numbering_definition` to
  build a new list style.
- :class:`NumberingDefinition` — wraps a ``w:abstractNum``. Call
  :meth:`NumberingDefinition.apply_to` to set a paragraph's numbering.
- :class:`Level` — read-only view of one level of a numbering definition,
  exposing ``number_format``, ``text``, and ``indent``.
"""

from __future__ import annotations

from collections import namedtuple
from typing import TYPE_CHECKING, Any, Iterator, List, Mapping, Sequence, Tuple, Union

from docx.enum.text import WD_NUMBER_FORMAT
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.shared import Length, Twips

if TYPE_CHECKING:
    from docx.oxml.numbering import CT_AbstractNum, CT_Lvl, CT_Numbering
    from docx.parts.numbering import NumberingPart
    from docx.text.paragraph import Paragraph


ListFormat = namedtuple("ListFormat", ("numbering_definition", "level"))
"""Named tuple returned by :attr:`Paragraph.list_format` for the read case.

``numbering_definition`` is the :class:`NumberingDefinition` (or |None| if the
paragraph is not part of a list), and ``level`` is the integer indent level
(``0`` through ``8``) or |None|.
"""


LevelSpec = Union[
    Mapping[str, Any],
    Sequence[Any],
]
"""A per-level specification. Accepts either a mapping with any of the keys
``format``, ``text``, ``start``, ``indent``, ``font`` or a positional tuple
``(format, text[, indent[, font]])``."""


def _normalize_format(value: Any) -> WD_NUMBER_FORMAT:
    """Return the :class:`WD_NUMBER_FORMAT` form for `value`.

    Accepts a :class:`WD_NUMBER_FORMAT` member or a raw XML string (e.g.
    ``"decimal"``).
    """
    if isinstance(value, WD_NUMBER_FORMAT):
        return value
    if isinstance(value, str):
        return WD_NUMBER_FORMAT.from_xml(value)
    raise TypeError(
        "format must be a WD_NUMBER_FORMAT member or string, got %r" % (value,)
    )


def _normalize_level_spec(
    spec: LevelSpec,
) -> Tuple[WD_NUMBER_FORMAT, str, int | None, str | None, int | None]:
    """Return ``(format, text, indent_twips, font, start)`` from `spec`.

    Missing values become |None|. ``indent_twips`` is an integer count of
    twentieths of a point (EMU-free for simplicity). ``font`` is the font name
    to apply via ``w:rPr/w:rFonts``.
    """
    if isinstance(spec, Mapping):
        fmt = _normalize_format(spec.get("format", "decimal"))
        text = spec.get("text", "%1.")
        indent = spec.get("indent")
        font = spec.get("font")
        start = spec.get("start")
    else:
        # -- positional tuple/list: (format, text[, indent[, font]]) --
        seq = list(spec)
        if len(seq) < 2:
            raise ValueError(
                "positional level spec must be (format, text[, indent[, font]])"
            )
        fmt = _normalize_format(seq[0])
        text = seq[1]
        indent = seq[2] if len(seq) > 2 else None
        font = seq[3] if len(seq) > 3 else None
        start = None

    if indent is not None and not isinstance(indent, int):
        # -- assume Length (subclass of int) via `Length` or accept bare int twips --
        indent = int(indent)  # pyright: ignore[reportGeneralTypeIssues]

    indent_twips: int | None
    if indent is None:
        indent_twips = None
    elif isinstance(indent, Length):
        # -- Length is EMU; convert to twips for w:ind --
        indent_twips = int(Length(indent).twips)
    else:
        # -- assume already twips --
        indent_twips = int(indent)

    return fmt, str(text), indent_twips, font, start


class Numbering:
    """Top-level proxy for a document's numbering part.

    Use ``document.numbering`` to obtain this object.
    """

    def __init__(self, numbering_elm: "CT_Numbering", part: "NumberingPart"):
        self._numbering = numbering_elm
        self._part = part

    @property
    def definitions(self) -> List["NumberingDefinition"]:
        """List of |NumberingDefinition| objects wrapping every ``w:abstractNum``."""
        return [
            NumberingDefinition(elm, self)
            for elm in self._numbering.abstractNum_lst
        ]

    def __iter__(self) -> Iterator["NumberingDefinition"]:
        return iter(self.definitions)

    def __len__(self) -> int:
        return len(self._numbering.abstractNum_lst)

    @property
    def element(self) -> "CT_Numbering":
        return self._numbering

    @property
    def part(self) -> "NumberingPart":
        return self._part

    def add_numbering_definition(
        self, levels: Sequence[LevelSpec]
    ) -> "NumberingDefinition":
        """Create and return a new |NumberingDefinition| built from `levels`.

        `levels` is a sequence of per-level specifications. Each element may be
        either a mapping::

            {"format": WD_NUMBER_FORMAT.DECIMAL, "text": "%1.", "indent": Inches(0.25)}

        or a positional tuple ``(format, text[, indent[, font]])``.

        `format` may be a :class:`WD_NUMBER_FORMAT` member or an OOXML string
        value (e.g. ``"decimal"`` or ``"bullet"``). `text` is the ``w:lvlText``
        pattern, using ``%N`` placeholders where ``N`` is 1-based. `indent`
        may be a :class:`~docx.shared.Length` or a raw twips integer. `font`
        sets the ``w:rFonts`` name on the level's run properties — required
        for bullet-style levels (e.g. ``"Symbol"`` for ``•``).
        """
        numbering = self._numbering
        abstractNum = numbering.add_abstractNum()

        # -- emit a w:multiLevelType hint matching the level count --
        multi_tag = (
            "singleLevel" if len(levels) == 1 else "hybridMultilevel"
        )
        multiLevelType = OxmlElement("w:multiLevelType")
        multiLevelType.set(qn("w:val"), multi_tag)
        abstractNum.append(multiLevelType)

        for ilvl, spec in enumerate(levels):
            fmt, text, indent_twips, font, start = _normalize_level_spec(spec)
            lvl = abstractNum.add_lvl()
            lvl.ilvl = ilvl
            lvl.start_val = start if start is not None else 1
            lvl.numFmt_val = fmt
            lvl.lvlText_val = text
            if indent_twips is not None:
                pPr = lvl.get_or_add_pPr()
                ind = pPr.get_or_add_ind()
                ind.left = Twips(indent_twips)
                # -- for numbered lists, a small hanging indent is typical --
                if fmt != WD_NUMBER_FORMAT.BULLET:
                    ind.hanging = Twips(360)
            if font is not None:
                rPr = lvl.get_or_add_rPr()
                rFonts = OxmlElement("w:rFonts")
                rFonts.set(qn("w:ascii"), font)
                rFonts.set(qn("w:hAnsi"), font)
                rFonts.set(qn("w:cs"), font)
                rPr.append(rFonts)

        # -- immediately create a matching w:num so the definition can be used --
        numbering.add_num(abstractNum.abstractNumId)

        return NumberingDefinition(abstractNum, self)

    def _num_id_for(self, abstractNum_id: int) -> int:
        """Return a ``numId`` pointing at `abstractNum_id`, reusing an existing
        ``w:num`` when possible, otherwise creating a new one."""
        for num in self._numbering.num_lst:
            abstractNumId = num.xpath("./w:abstractNumId")
            if not abstractNumId:
                continue
            if int(abstractNumId[0].get(qn("w:val"))) == abstractNum_id:
                return num.numId
        new_num = self._numbering.add_num(abstractNum_id)
        return new_num.numId


class NumberingDefinition:
    """Proxy for a single ``w:abstractNum`` element."""

    def __init__(
        self,
        abstractNum: "CT_AbstractNum",
        numbering: "Numbering",
    ):
        self._abstractNum = abstractNum
        self._numbering = numbering

    @property
    def abstract_num_id(self) -> int:
        """The integer id of this abstract numbering definition."""
        return self._abstractNum.abstractNumId

    @property
    def element(self) -> "CT_AbstractNum":
        return self._abstractNum

    @property
    def levels(self) -> List["Level"]:
        """List of :class:`Level` objects, one per declared ``w:lvl``."""
        return [Level(lvl, self) for lvl in self._abstractNum.lvl_lst]

    def level(self, ilvl: int) -> "Level | None":
        """Return the |Level| with `ilvl`, or |None| if none exists."""
        lvl = self._abstractNum.get_lvl(ilvl)
        if lvl is None:
            return None
        return Level(lvl, self)

    def apply_to(self, paragraph: "Paragraph", level: int = 0) -> None:
        """Apply this numbering definition to `paragraph` at the specified `level`.

        Sets the paragraph's ``w:numPr`` children ``w:numId`` (resolving a
        matching ``w:num`` instance, creating one if necessary) and ``w:ilvl``.
        """
        if not 0 <= level <= 8:
            raise ValueError("level must be in range 0..8, got %d" % level)
        num_id = self._numbering._num_id_for(self.abstract_num_id)
        p = paragraph._p  # pyright: ignore[reportPrivateUsage]
        pPr = p.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        numPr.ilvl_val = level
        numPr.numId_val = num_id


class Level:
    """Read-only view of one ``w:lvl`` child of a ``w:abstractNum``."""

    def __init__(self, lvl: "CT_Lvl", definition: "NumberingDefinition"):
        self._lvl = lvl
        self._definition = definition

    @property
    def ilvl(self) -> int:
        """Zero-based indent level of this level."""
        return self._lvl.ilvl

    @property
    def number_format(self) -> WD_NUMBER_FORMAT | None:
        """The :class:`WD_NUMBER_FORMAT` member corresponding to ``w:numFmt/@val``.

        Returns |None| if no ``w:numFmt`` is present, or if the XML value is
        outside the subset of formats exposed by :class:`WD_NUMBER_FORMAT`.
        """
        try:
            return self._lvl.numFmt_val
        except ValueError:
            return None

    @property
    def text(self) -> str | None:
        """The ``w:lvlText/@val`` pattern, e.g. ``"%1."`` or ``"%1.%2"``."""
        return self._lvl.lvlText_val

    @property
    def start(self) -> int:
        """The starting value (``w:start/@val``), defaulting to ``1``."""
        return self._lvl.start_val

    @property
    def indent(self) -> Length | None:
        """The ``w:left`` indent declared on this level, or |None|."""
        pPr = self._lvl.pPr
        if pPr is None:
            return None
        ind = pPr.ind
        if ind is None:
            return None
        return ind.left

    @property
    def element(self) -> "CT_Lvl":
        return self._lvl
