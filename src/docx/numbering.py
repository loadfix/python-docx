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

import re
from collections import namedtuple
from typing import TYPE_CHECKING, Any, Union
from collections.abc import Iterator, Mapping, Sequence

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_NUMBER_FORMAT
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.shared import Length, Twips

if TYPE_CHECKING:
    from docx.oxml.numbering import (
        CT_AbstractNum,
        CT_Lvl,
        CT_Num,
        CT_NumLvl,
        CT_Numbering,
        CT_NumPicBullet,
    )
    from docx.oxml.text.paragraph import CT_P
    from docx.parts.numbering import NumberingPart
    from docx.text.paragraph import Paragraph


ListFormat = namedtuple("ListFormat", ("numbering_definition", "level"))
"""Named tuple returned by :attr:`Paragraph.list_format` for the read case.

``numbering_definition`` is the :class:`NumberingDefinition` (or |None| if the
paragraph is not part of a list), and ``level`` is the integer indent level
(``0`` through ``8``) or |None|.
"""


LevelSpec = Union[Mapping[str, Any], Sequence[Any]]
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
) -> tuple[WD_NUMBER_FORMAT, str, int | None, str | None, int | None]:
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

    .. versionadded:: 2026.05.0
    """

    def __init__(self, numbering_elm: "CT_Numbering", part: "NumberingPart"):
        self._numbering = numbering_elm
        self._part = part

    @property
    def picture_bullets(self) -> list["PictureBullet"]:
        """List of |PictureBullet| proxies wrapping each ``w:numPicBullet``.

        Word authors ``<w:numPicBullet>`` entries when the user picks
        *Home* > *Bullets* > *Define New Bullet* > *Picture*. Each bullet is
        identified by a stable ``numPicBulletId`` that level definitions
        reference via ``<w:numPicBulletId w:val="..."/>``.

        .. versionadded:: 2026.05.3
        """
        return [PictureBullet(elm, self) for elm in self._numbering.numPicBullet_lst]

    def picture_bullet(self, numPicBulletId: int) -> "PictureBullet | None":
        """Return the |PictureBullet| with matching ``numPicBulletId``, or |None|.

        .. versionadded:: 2026.05.3
        """
        for elm in self._numbering.numPicBullet_lst:
            if elm.numPicBulletId == numPicBulletId:
                return PictureBullet(elm, self)
        return None

    def remove_picture_bullet(self, numPicBulletId: int) -> bool:
        """Remove the ``w:numPicBullet`` whose id matches `numPicBulletId`.

        Returns |True| when a matching element was removed, |False| otherwise.

        .. versionadded:: 2026.05.3
        """
        for elm in list(self._numbering.numPicBullet_lst):
            if elm.numPicBulletId == numPicBulletId:
                self._numbering.remove(elm)
                return True
        return False

    @property
    def _next_numPicBulletId(self) -> int:
        """The lowest positive integer not already used by a ``w:numPicBullet``.

        .. versionadded:: 2026.05.3
        """
        used = {elm.numPicBulletId for elm in self._numbering.numPicBullet_lst}
        candidate = 0
        while candidate in used:
            candidate += 1
        return candidate

    @property
    def definitions(self) -> list["NumberingDefinition"]:
        """List of |NumberingDefinition| objects wrapping every ``w:abstractNum``.

        .. versionadded:: 2026.05.0
        """
        return [
            NumberingDefinition(elm, self)
            for elm in self._numbering.abstractNum_lst
        ]

    @property
    def abstract_definitions(self) -> list["NumberingDefinition"]:
        """Alias of :attr:`definitions` — the list of abstract numbering definitions.

        Mirrors the ``abstractNum`` vs ``num`` split in WordprocessingML so that
        callers writing new code can address abstract definitions explicitly.

        .. versionadded:: 2026.05.10
        """
        return self.definitions

    @property
    def num_instances(self) -> list["NumInstance"]:
        """List of |NumInstance| proxies wrapping every ``w:num`` instance.

        A ``w:num`` is a concrete instance of an abstract numbering definition
        and carries the ``numId`` referenced from a paragraph's ``w:numPr``.

        .. versionadded:: 2026.05.10
        """
        return [NumInstance(elm, self) for elm in self._numbering.num_lst]

    def num_instance(self, num_id: int) -> "NumInstance | None":
        """Return the |NumInstance| with matching `num_id`, or |None|.

        .. versionadded:: 2026.05.10
        """
        try:
            elm = self._numbering.num_having_numId(num_id)
        except KeyError:
            return None
        return NumInstance(elm, self)

    def abstract_definition(
        self, abstract_num_id: int
    ) -> "NumberingDefinition | None":
        """Return the |NumberingDefinition| with matching `abstract_num_id`, or |None|.

        .. versionadded:: 2026.05.10
        """
        try:
            elm = self._numbering.abstractNum_having_abstractNumId(abstract_num_id)
        except KeyError:
            return None
        return NumberingDefinition(elm, self)

    def next_num_id(self) -> int:
        """The next unused ``numId`` that :meth:`add_definition` would allocate.

        Peeking does not create a ``w:num`` element; it simply reports what
        :attr:`CT_Numbering.next_numId` currently is.

        .. versionadded:: 2026.05.10
        """
        return self._numbering.next_numId

    def next_abstract_num_id(self) -> int:
        """The next unused ``abstractNumId`` that :meth:`add_abstract_definition`
        would allocate.

        .. versionadded:: 2026.05.10
        """
        return self._numbering.next_abstractNumId

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

        .. versionadded:: 2026.05.0
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

    def add_abstract_definition(
        self,
        format: "WD_NUMBER_FORMAT | str" = WD_NUMBER_FORMAT.DECIMAL,
        start: int = 1,
        lvl_text: str = "%1.",
        alignment: "WD_ALIGN_PARAGRAPH | None" = WD_ALIGN_PARAGRAPH.LEFT,
    ) -> "NumberingDefinition":
        """Create and return a single-level abstract numbering definition.

        This is a convenience constructor for the common *"I just want a
        decimal list"* case. It creates a ``w:abstractNum`` with one level
        (``ilvl="0"``) configured with `format`, `start`, `lvl_text`, and
        `alignment`. For multi-level lists use
        :meth:`add_numbering_definition` which accepts a per-level spec.

        Note that no ``w:num`` instance is created; call
        :meth:`add_definition` with the returned definition's
        ``abstract_num_id`` (or :meth:`NumberingDefinition.new_instance`) to
        allocate a usable ``numId``.

        .. versionadded:: 2026.05.10
        """
        fmt = _normalize_format(format)
        numbering = self._numbering
        abstractNum = numbering.add_abstractNum()
        multiLevelType = OxmlElement("w:multiLevelType")
        multiLevelType.set(qn("w:val"), "singleLevel")
        abstractNum.append(multiLevelType)

        lvl = abstractNum.add_lvl()
        lvl.ilvl = 0
        lvl.start_val = int(start)
        lvl.numFmt_val = fmt
        lvl.lvlText_val = lvl_text
        if alignment is not None:
            lvl.lvlJc_val = alignment.xml_value

        return NumberingDefinition(abstractNum, self)

    def add_definition(
        self, abstract_num_id: int, style_name: str | None = None
    ) -> "NumInstance":
        """Allocate a new ``w:num`` instance referencing `abstract_num_id`.

        Returns a |NumInstance| wrapping the new element. When `style_name`
        is given, a ``w:pStyle`` child is attached to the abstract definition's
        level-0 so Word will associate the style with the list.

        Raises ``KeyError`` if no ``w:abstractNum`` with `abstract_num_id`
        exists in the numbering part.

        .. versionadded:: 2026.05.10
        """
        # -- verify abstract_num_id resolves before mutating --
        abstractNum = self._numbering.abstractNum_having_abstractNumId(
            abstract_num_id
        )
        num = self._numbering.add_num(abstract_num_id)
        if style_name is not None:
            lvl0 = abstractNum.get_lvl(0)
            if lvl0 is None:
                lvl0 = abstractNum.add_lvl()
                lvl0.ilvl = 0
            lvl0.pStyle_val = style_name
        return NumInstance(num, self)

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
    """Proxy for a single ``w:abstractNum`` element.

    .. versionadded:: 2026.05.0
    """

    def __init__(
        self,
        abstractNum: "CT_AbstractNum",
        numbering: "Numbering",
    ):
        self._abstractNum = abstractNum
        self._numbering = numbering

    @property
    def abstract_num_id(self) -> int:
        """The integer id of this abstract numbering definition.

        .. versionadded:: 2026.05.0
        """
        return self._abstractNum.abstractNumId

    @property
    def element(self) -> "CT_AbstractNum":
        return self._abstractNum

    @property
    def levels(self) -> list["Level"]:
        """List of :class:`Level` objects, one per declared ``w:lvl``.

        .. versionadded:: 2026.05.0
        """
        return [Level(lvl, self) for lvl in self._abstractNum.lvl_lst]

    def level(self, ilvl: int) -> "Level | None":
        """Return the |Level| with `ilvl`, or |None| if none exists.

        .. versionadded:: 2026.05.0
        """
        lvl = self._abstractNum.get_lvl(ilvl)
        if lvl is None:
            return None
        return Level(lvl, self)

    def new_instance(self) -> int:
        """Allocate a new ``w:num`` pointing at this abstract definition.

        Returns the integer ``numId`` of the freshly-created instance. Two
        paragraphs sharing the same abstract definition but different
        ``numId`` values restart their numbering independently; this helper
        is handy for laying out several independent lists that all share the
        same visual formatting (closes upstream#25).

        .. versionadded:: 2026.05.0
        """
        new_num = self._numbering._numbering.add_num(self.abstract_num_id)
        return new_num.numId

    def apply_to(self, paragraph: "Paragraph", level: int = 0) -> None:
        """Apply this numbering definition to `paragraph` at the specified `level`.

        Sets the paragraph's ``w:numPr`` children ``w:numId`` (resolving a
        matching ``w:num`` instance, creating one if necessary) and ``w:ilvl``.

        .. versionadded:: 2026.05.0
        """
        if not 0 <= level <= 8:
            raise ValueError("level must be in range 0..8, got %d" % level)
        num_id = self._numbering._num_id_for(self.abstract_num_id)
        p = paragraph._p  # pyright: ignore[reportPrivateUsage]
        pPr = p.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        numPr.ilvl_val = level
        numPr.numId_val = num_id


#: Alias of :class:`NumberingDefinition` emphasising that each object wraps a
#: ``w:abstractNum`` element.  New code may prefer the longer name for clarity;
#: existing ``NumberingDefinition`` callers continue to work unchanged.
#:
#: .. versionadded:: 2026.05.10
AbstractNumberingDefinition = NumberingDefinition


class NumInstance:
    """Proxy for a single ``w:num`` element — a concrete numbering instance.

    A ``w:num`` carries the ``numId`` that paragraphs reference in their
    ``w:numPr/w:numId`` and points at an abstract definition. Instances are
    created implicitly by :meth:`NumberingDefinition.new_instance` or
    explicitly by :meth:`Numbering.add_definition`.

    .. versionadded:: 2026.05.10
    """

    def __init__(self, num: "CT_Num", numbering: "Numbering"):
        self._num = num
        self._numbering = numbering

    @property
    def num_id(self) -> int:
        """The ``numId`` referenced from paragraphs' ``w:numPr/w:numId``."""
        return self._num.numId

    @property
    def abstract_num_id(self) -> int:
        """The ``abstractNumId`` of the definition this instance points at."""
        return self._num.abstractNumId_val

    @property
    def definition(self) -> "NumberingDefinition | None":
        """The |NumberingDefinition| this instance references, or |None|.

        Returns |None| if the numbering part has no matching ``w:abstractNum``
        (a malformed document).
        """
        return self._numbering.abstract_definition(self.abstract_num_id)

    @property
    def level_overrides(self) -> list["LevelOverride"]:
        """List of |LevelOverride| proxies for each ``w:lvlOverride`` child."""
        return [LevelOverride(elm, self) for elm in self._num.lvlOverride_lst]

    def level_override(self, ilvl: int) -> "LevelOverride | None":
        """Return the |LevelOverride| at `ilvl`, or |None| when absent."""
        elm = self._num.get_lvlOverride(ilvl)
        if elm is None:
            return None
        return LevelOverride(elm, self)

    def set_level_override(
        self, ilvl: int, start: int | None = None
    ) -> "LevelOverride":
        """Create (or return) the ``w:lvlOverride`` for `ilvl` and optionally set
        its ``w:startOverride`` to `start`.

        Returns the |LevelOverride| proxy. When `start` is |None| the override
        element is still created but no ``w:startOverride`` child is written —
        callers can add level-level formatting later.
        """
        if not 0 <= ilvl <= 8:
            raise ValueError("ilvl must be in range 0..8, got %d" % ilvl)
        override = self._num.get_or_add_lvlOverride(ilvl)
        if start is not None:
            # -- remove any pre-existing startOverride so repeat calls stomp --
            for child in override.xpath("./w:startOverride"):
                override.remove(child)
            override.add_startOverride(val=int(start))
        return LevelOverride(override, self)

    @property
    def element(self) -> "CT_Num":
        return self._num


class LevelOverride:
    """Proxy for a ``w:lvlOverride`` element on a ``w:num`` instance.

    Level overrides let a single list instance redefine the starting value or
    (optionally) swap an entire level of formatting compared to the underlying
    abstract definition.  This proxy exposes the common ``startOverride`` lever;
    the underlying ``w:lvl`` child remains available via :attr:`element`.

    .. versionadded:: 2026.05.10
    """

    def __init__(self, override: "CT_NumLvl", num_instance: "NumInstance"):
        self._override = override
        self._num_instance = num_instance

    @property
    def ilvl(self) -> int:
        """Zero-based level index this override targets."""
        return self._override.ilvl

    @property
    def start_override(self) -> int | None:
        """The override starting value, or |None| when no ``w:startOverride`` is set."""
        elems = self._override.xpath("./w:startOverride")
        if not elems:
            return None
        val = elems[0].get(qn("w:val"))
        return int(val) if val is not None else None

    @start_override.setter
    def start_override(self, value: int | None) -> None:
        for child in self._override.xpath("./w:startOverride"):
            self._override.remove(child)
        if value is not None:
            self._override.add_startOverride(val=int(value))

    @property
    def element(self) -> "CT_NumLvl":
        return self._override


class Level:
    """Read-only view of one ``w:lvl`` child of a ``w:abstractNum``.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, lvl: "CT_Lvl", definition: "NumberingDefinition"):
        self._lvl = lvl
        self._definition = definition

    @property
    def ilvl(self) -> int:
        """Zero-based indent level of this level.

        .. versionadded:: 2026.05.0
        """
        return self._lvl.ilvl

    @property
    def number_format(self) -> WD_NUMBER_FORMAT | None:
        """The :class:`WD_NUMBER_FORMAT` member corresponding to ``w:numFmt/@val``.

        Returns |None| if no ``w:numFmt`` is present, or if the XML value is
        outside the subset of formats exposed by :class:`WD_NUMBER_FORMAT`.

        .. versionadded:: 2026.05.0
        """
        try:
            return self._lvl.numFmt_val
        except ValueError:
            return None

    @property
    def text(self) -> str | None:
        """The ``w:lvlText/@val`` pattern, e.g. ``"%1."`` or ``"%1.%2"``.

        .. versionadded:: 2026.05.0
        """
        return self._lvl.lvlText_val

    @property
    def start(self) -> int:
        """The starting value (``w:start/@val``), defaulting to ``1``.

        .. versionadded:: 2026.05.0
        """
        return self._lvl.start_val

    @property
    def indent(self) -> Length | None:
        """The ``w:left`` indent declared on this level, or |None|.

        .. versionadded:: 2026.05.0
        """
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


# -- lvlText placeholder parser: matches %1, %2, ... through %9 --
_LVLTEXT_TOKEN_RE = re.compile(r"%([1-9])")


def _format_decimal(n: int) -> str:
    return str(n)


def _format_decimal_zero(n: int) -> str:
    # -- Word renders single digits with a leading zero (01, 02, ..., 09, 10, 11) --
    return "%02d" % n if n < 10 else str(n)


_ROMAN_PAIRS = (
    (1000, "M"),
    (900, "CM"),
    (500, "D"),
    (400, "CD"),
    (100, "C"),
    (90, "XC"),
    (50, "L"),
    (40, "XL"),
    (10, "X"),
    (9, "IX"),
    (5, "V"),
    (4, "IV"),
    (1, "I"),
)


def _format_upper_roman(n: int) -> str:
    if n <= 0:
        return str(n)
    out: list[str] = []
    remaining = n
    for value, numeral in _ROMAN_PAIRS:
        while remaining >= value:
            out.append(numeral)
            remaining -= value
    return "".join(out)


def _format_lower_roman(n: int) -> str:
    return _format_upper_roman(n).lower()


def _format_letter(n: int, base: str) -> str:
    """Return the ``n``-th letter sequence using `base` as the starting letter.

    Word's ``lowerLetter`` / ``upperLetter`` use spreadsheet-column semantics:
    1->a, 2->b, ..., 26->z, 27->aa, 28->ab, ..., 52->az, 53->ba, etc.
    """
    if n <= 0:
        return str(n)
    # -- treat as base-26 with no zero digit; convert via the classic "while n > 0" --
    result = ""
    remaining = n
    base_ord = ord(base)
    while remaining > 0:
        remaining, rem = divmod(remaining - 1, 26)
        result = chr(base_ord + rem) + result
    return result


def _format_upper_letter(n: int) -> str:
    return _format_letter(n, "A")


def _format_lower_letter(n: int) -> str:
    return _format_letter(n, "a")


_NUMFMT_FORMATTERS: dict[WD_NUMBER_FORMAT, Any] = {
    WD_NUMBER_FORMAT.DECIMAL: _format_decimal,
    WD_NUMBER_FORMAT.UPPER_ROMAN: _format_upper_roman,
    WD_NUMBER_FORMAT.LOWER_ROMAN: _format_lower_roman,
    WD_NUMBER_FORMAT.UPPER_LETTER: _format_upper_letter,
    WD_NUMBER_FORMAT.LOWER_LETTER: _format_lower_letter,
}
# -- decimalZero isn't exposed on WD_NUMBER_FORMAT; handled by raw-XML lookup. --


def _format_counter(fmt: WD_NUMBER_FORMAT | None, raw_fmt: str | None, n: int) -> str:
    """Return the rendered string for counter `n` given its format.

    `fmt` is the :class:`WD_NUMBER_FORMAT` member (or |None| when the XML
    ``w:numFmt/@val`` isn't in our enum subset); `raw_fmt` is the raw
    ``w:numFmt/@val`` string — used to pick up formats the enum doesn't
    cover (e.g. ``decimalZero``).
    """
    if raw_fmt == "decimalZero":
        return _format_decimal_zero(n)
    if fmt is None:
        # -- unknown format: fall back to decimal --
        return _format_decimal(n)
    formatter = _NUMFMT_FORMATTERS.get(fmt)
    if formatter is not None:
        return formatter(n)
    # TODO: cardinalText / ordinalText / ordinal / chicago — fall back to decimal.
    return _format_decimal(n)


class ListLabelRenderer:
    """Stateful walker that renders the Word-style label for each numbered paragraph.

    A new renderer is created per document (or per traversal). It walks the
    document body in order, maintaining per-``abstractNum`` counters keyed by
    level, and produces the rendered label string (``"1."``, ``"a)"``,
    ``"•"``, etc.) for each paragraph.

    The class is intentionally standalone — it accepts the ``CT_Numbering``
    element and paragraph elements directly, so it can be used both from
    :attr:`Paragraph.list_label` (lazily, for a single paragraph) and from
    :meth:`Document.list_labels` (eagerly, for every paragraph).

    .. versionadded:: 2026.05.0
    """

    def __init__(
        self,
        numbering_elm: "CT_Numbering | None",
        styles_elm: Any | None = None,
    ):
        self._numbering = numbering_elm
        self._styles = styles_elm
        # -- per-abstractNum counter state: {abstract_num_id: {ilvl: count}} --
        self._counters: dict[int, dict[int, int]] = {}
        # -- cached numId -> abstractNumId resolution --
        self._num_to_abstract: dict[int, int] = {}
        # -- cached abstractNum element lookup --
        self._abstract_cache: dict[int, "CT_AbstractNum | None"] = {}
        # -- cached style numId inheritance: {styleId: (numId, ilvl)} --
        self._style_num_cache: dict[str, tuple[int | None, int | None]] = {}
        # -- cached rendered labels keyed by paragraph element id (for repeat lookups) --
        self._rendered: dict[int, str | None] = {}

    def label_for(self, p: "CT_P") -> str | None:
        """Return the rendered list label for paragraph element `p`, or |None|.

        Returns |None| when `p` is not part of any numbered list (no
        ``w:numPr/w:numId`` directly or via its style), when the referenced
        ``w:num``/``w:abstractNum`` cannot be resolved, or when the level's
        ``w:lvlText`` is absent.

        Each call advances the per-abstract counter state, so callers must
        invoke ``label_for`` for every paragraph *in document order* or use
        :meth:`label_map` which does that for them.
        """
        key = id(p)
        if key in self._rendered:
            return self._rendered[key]
        label = self._compute_label(p)
        self._rendered[key] = label
        return label

    def label_map(self, paragraphs: Iterator["CT_P"]) -> dict[int, str]:
        """Walk `paragraphs` in order and return ``{id(p): label}`` for labelled paragraphs.

        Only paragraphs that resolve to a non-None label are included.
        """
        result: dict[int, str] = {}
        for p in paragraphs:
            label = self.label_for(p)
            if label is not None:
                result[id(p)] = label
        return result

    # -- internal helpers ------------------------------------------------

    def _compute_label(self, p: "CT_P") -> str | None:
        num_id, ilvl = self._resolve_numPr(p)
        if num_id is None or ilvl is None:
            return None
        if self._numbering is None:
            return None

        abstract_id = self._abstract_num_id_for(num_id)
        if abstract_id is None:
            return None

        abstractNum = self._get_abstractNum(abstract_id)
        if abstractNum is None:
            return None

        lvl = abstractNum.get_lvl(ilvl)
        if lvl is None:
            return None

        # -- bump counter for current level and reset deeper levels --
        self._advance_counter(abstract_id, ilvl, abstractNum)

        # -- render the lvlText, substituting %N with the appropriate counter --
        lvlText = lvl.lvlText_val
        if lvlText is None:
            return None

        return self._render_lvlText(lvlText, abstract_id, abstractNum)

    def _resolve_numPr(self, p: "CT_P") -> tuple[int | None, int | None]:
        """Return ``(numId, ilvl)`` for paragraph `p`, consulting its style if necessary."""
        pPr = p.pPr
        num_id: int | None = None
        ilvl: int | None = None
        if pPr is not None and pPr.numPr is not None:
            num_id = pPr.numPr.numId_val
            ilvl = pPr.numPr.ilvl_val
        if num_id is not None:
            if ilvl is None:
                ilvl = 0
            return num_id, ilvl
        # -- direct numPr absent or missing numId: consult style chain --
        style_id = None
        if pPr is not None:
            pStyle = pPr.style
            if pStyle is not None:
                style_id = pStyle
        if style_id is None:
            return None, None
        style_num = self._numPr_from_style(style_id)
        if style_num is None:
            return None, None
        s_num_id, s_ilvl = style_num
        if s_num_id is None:
            return None, None
        # -- direct ilvl (if any) wins over style-inherited --
        return s_num_id, ilvl if ilvl is not None else (s_ilvl or 0)

    def _numPr_from_style(self, style_id: str) -> tuple[int | None, int | None] | None:
        """Walk style → basedOn chain looking for a w:pPr/w:numPr.

        Returns ``(numId, ilvl)`` or |None| when no style in the chain declares
        numbering. Uses a seen-set to guard against a cyclic basedOn chain.
        """
        if self._styles is None:
            return None
        if style_id in self._style_num_cache:
            return self._style_num_cache[style_id]

        seen: set[str] = set()
        cur = style_id
        while cur is not None and cur not in seen:
            seen.add(cur)
            style = self._styles.get_by_id(cur)
            if style is None:
                break
            pPr = getattr(style, "pPr", None)
            if pPr is not None and pPr.numPr is not None:
                num_id = pPr.numPr.numId_val
                ilvl = pPr.numPr.ilvl_val
                self._style_num_cache[style_id] = (num_id, ilvl)
                return num_id, ilvl
            basedOn = style.basedOn_val
            if basedOn is None:
                break
            cur = basedOn
        self._style_num_cache[style_id] = (None, None)
        return None

    def _abstract_num_id_for(self, num_id: int) -> int | None:
        """Resolve ``w:num`` → ``w:abstractNumId`` for `num_id`."""
        if num_id in self._num_to_abstract:
            return self._num_to_abstract[num_id]
        assert self._numbering is not None
        try:
            num = self._numbering.num_having_numId(num_id)
        except KeyError:
            return None
        try:
            abstract_id = num.abstractNumId.val
        except AttributeError:
            return None
        self._num_to_abstract[num_id] = abstract_id
        return abstract_id

    def _get_abstractNum(self, abstract_id: int) -> "CT_AbstractNum | None":
        if abstract_id in self._abstract_cache:
            return self._abstract_cache[abstract_id]
        assert self._numbering is not None
        try:
            abstractNum = self._numbering.abstractNum_having_abstractNumId(abstract_id)
        except KeyError:
            abstractNum = None
        self._abstract_cache[abstract_id] = abstractNum
        return abstractNum

    def _advance_counter(
        self, abstract_id: int, ilvl: int, abstractNum: "CT_AbstractNum"
    ) -> None:
        """Increment the counter at `ilvl` and reset deeper levels to their start.

        The start value for each level is read from the level's ``w:start``
        value (defaulting to 1). When the paragraph's level is deeper than
        any level previously visited in this abstractNum, all shallower
        counters remain at their current value.
        """
        counters = self._counters.setdefault(abstract_id, {})
        # -- reset any deeper levels so a subsequent return to a deeper level restarts --
        for deeper_ilvl in list(counters.keys()):
            if deeper_ilvl > ilvl:
                del counters[deeper_ilvl]
        # -- increment (or initialise) the current level --
        if ilvl in counters:
            counters[ilvl] += 1
        else:
            lvl = abstractNum.get_lvl(ilvl)
            start = lvl.start_val if lvl is not None else 1
            counters[ilvl] = start

    def _render_lvlText(
        self, lvlText: str, abstract_id: int, abstractNum: "CT_AbstractNum"
    ) -> str:
        """Substitute ``%N`` placeholders in `lvlText` with formatted counters."""
        counters = self._counters.get(abstract_id, {})

        def replace(match: re.Match[str]) -> str:
            level_index = int(match.group(1)) - 1  # %1 is level 0
            n = counters.get(level_index)
            if n is None:
                # -- counter for a level we haven't entered yet (rare but possible
                # -- when a list starts at a deeper level): fall back to the
                # -- declared start value for that level.
                lvl = abstractNum.get_lvl(level_index)
                n = lvl.start_val if lvl is not None else 1
            lvl = abstractNum.get_lvl(level_index)
            if lvl is None:
                return _format_decimal(n)
            raw_fmt = None
            numFmt_elm = lvl.numFmt
            if numFmt_elm is not None:
                raw_fmt = numFmt_elm.get(qn("w:val"))
            try:
                fmt = lvl.numFmt_val
            except ValueError:
                # -- w:numFmt/@val is outside the WD_NUMBER_FORMAT subset
                # -- (e.g. ``decimalZero``); rely on ``raw_fmt`` lookup below --
                fmt = None
            # -- bullet-in-lvlText is a malformed case; emit empty so the
            # -- surrounding verbatim text is preserved --
            if fmt == WD_NUMBER_FORMAT.BULLET:
                return ""
            return _format_counter(fmt, raw_fmt, n)

        # -- short-circuit: if the lvlText contains no %N tokens (the common
        # -- case for bullets), return it verbatim. --
        if not _LVLTEXT_TOKEN_RE.search(lvlText):
            return lvlText
        return _LVLTEXT_TOKEN_RE.sub(replace, lvlText)


class PictureBullet:
    """Proxy for a single ``w:numPicBullet`` element in numbering.xml.

    Picture bullets are created via *Home* > *Bullets* > *Define New Bullet*
    > *Picture* in Word. The bullet image is an ordinary ``w:drawing`` child
    referencing an image part in the document's relationships; the picture is
    identified within numbering.xml by its ``@w:numPicBulletId`` so list-level
    definitions can cross-reference it.

    .. versionadded:: 2026.05.3
    """

    def __init__(self, element: "CT_NumPicBullet", numbering: "Numbering"):
        self._element = element
        self._numbering = numbering

    @property
    def id(self) -> int:
        """The ``@w:numPicBulletId`` of this bullet entry.

        .. versionadded:: 2026.05.3
        """
        return self._element.numPicBulletId

    @property
    def drawing(self):
        """The inner ``w:drawing`` element, or |None| if absent.

        The returned value is the raw ``CT_Drawing`` element (same type used
        by ``InlineShape``) for callers that need to inspect or replace the
        picture payload.

        .. versionadded:: 2026.05.3
        """
        return self._element.drawing

    @property
    def element(self) -> "CT_NumPicBullet":
        """The underlying ``w:numPicBullet`` element.

        .. versionadded:: 2026.05.3
        """
        return self._element
