"""Custom element classes related to the numbering part."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast
from collections.abc import Callable

from docx.oxml.ns import nsdecls
from docx.oxml.parser import OxmlElement, parse_xml
from docx.oxml.shared import CT_DecimalNumber
from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from docx.enum.text import WD_NUMBER_FORMAT
    from docx.oxml.footnotes import CT_NumFmt
    from docx.oxml.text.parfmt import CT_PPr
    from docx.oxml.text.font import CT_RPr


class CT_LvlText(BaseOxmlElement):
    """``<w:lvlText>`` element, holding the format pattern for a numbering level."""

    val: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_Lvl(BaseOxmlElement):
    """``<w:lvl>`` element, one of up to nine levels defining a list's visual
    formatting inside an abstract numbering definition."""

    get_or_add_pPr: Callable[[], "CT_PPr"]
    get_or_add_rPr: Callable[[], "CT_RPr"]

    _tag_seq = (
        "w:start",
        "w:numFmt",
        "w:lvlRestart",
        "w:pStyle",
        "w:isLgl",
        "w:suff",
        "w:lvlText",
        "w:lvlPicBulletId",
        "w:legacy",
        "w:lvlJc",
        "w:pPr",
        "w:rPr",
    )

    start: CT_DecimalNumber | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:start", successors=_tag_seq[1:]
    )
    numFmt: CT_NumFmt | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:numFmt", successors=_tag_seq[2:]
    )
    lvlText: CT_LvlText | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:lvlText", successors=_tag_seq[7:]
    )
    lvlJc = ZeroOrOne("w:lvlJc", successors=_tag_seq[10:])
    pPr = ZeroOrOne("w:pPr", successors=_tag_seq[11:])
    rPr = ZeroOrOne("w:rPr", successors=())

    ilvl: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:ilvl", ST_DecimalNumber
    )

    del _tag_seq

    @property
    def start_val(self) -> int:
        """The integer ``val`` attribute of the ``<w:start>`` grandchild.

        Returns ``1`` if no ``<w:start>`` child is present.
        """
        start = self.start
        if start is None:
            return 1
        return start.val

    @start_val.setter
    def start_val(self, value: int) -> None:
        start = self.get_or_add_start()
        start.val = value

    @property
    def numFmt_val(self) -> "WD_NUMBER_FORMAT | None":
        """The ``val`` attribute of ``<w:numFmt>`` or |None| if absent."""
        numFmt = self.numFmt
        if numFmt is None:
            return None
        return numFmt.val

    @numFmt_val.setter
    def numFmt_val(self, value: "WD_NUMBER_FORMAT | None") -> None:
        numFmt = self.get_or_add_numFmt()
        numFmt.val = value

    @property
    def lvlText_val(self) -> str | None:
        """The ``val`` attribute of ``<w:lvlText>`` or |None| if absent."""
        lvlText = self.lvlText
        if lvlText is None:
            return None
        return lvlText.val

    @lvlText_val.setter
    def lvlText_val(self, value: str) -> None:
        lvlText = self.get_or_add_lvlText()
        lvlText.val = value


class CT_AbstractNum(BaseOxmlElement):
    """``<w:abstractNum>`` element, defining an abstract numbering definition.

    Holds up to nine ``<w:lvl>`` children, one per list level.
    """

    lvl_lst: list[CT_Lvl]
    add_lvl: Callable[..., CT_Lvl]

    lvl = ZeroOrMore("w:lvl", successors=())

    abstractNumId: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:abstractNumId", ST_DecimalNumber
    )

    def get_lvl(self, ilvl: int) -> CT_Lvl | None:
        """Return the ``<w:lvl>`` child with matching `ilvl`, or |None|."""
        for lvl in self.lvl_lst:
            if lvl.ilvl == ilvl:
                return lvl
        return None


class CT_NumPicBullet(BaseOxmlElement):
    """``<w:numPicBullet>`` element, a picture used as a custom bullet glyph.

    Contains a single ``<w:drawing>`` child wrapping the bullet image plus a
    required ``@w:numPicBulletId`` attribute that level definitions reference
    via ``<w:numPicBulletId w:val="..."/>`` to use the picture as the bullet
    for that level.

    Created by Word's *Home* > *Bullets* > *Define New Bullet* > *Picture*
    command.

    .. versionadded:: 2026.05.3
    """

    drawing = ZeroOrOne("w:drawing", successors=())
    numPicBulletId: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:numPicBulletId", ST_DecimalNumber
    )


class CT_Num(BaseOxmlElement):
    """``<w:num>`` element, which represents a concrete list definition instance, having
    a required child <w:abstractNumId> that references an abstract numbering definition
    that defines most of the formatting details."""

    abstractNumId = OneAndOnlyOne("w:abstractNumId")
    lvlOverride = ZeroOrMore("w:lvlOverride")
    numId = RequiredAttribute("w:numId", ST_DecimalNumber)

    def add_lvlOverride(self, ilvl):
        """Return a newly added CT_NumLvl (<w:lvlOverride>) element having its ``ilvl``
        attribute set to `ilvl`."""
        return self._add_lvlOverride(ilvl=ilvl)

    @classmethod
    def new(cls, num_id, abstractNum_id):
        """Return a new ``<w:num>`` element having numId of `num_id` and having a
        ``<w:abstractNumId>`` child with val attribute set to `abstractNum_id`."""
        num = OxmlElement("w:num")
        num.numId = num_id
        abstractNumId = CT_DecimalNumber.new("w:abstractNumId", abstractNum_id)
        num.append(abstractNumId)
        return num


class CT_NumLvl(BaseOxmlElement):
    """``<w:lvlOverride>`` element, which identifies a level in a list definition to
    override with settings it contains."""

    startOverride = ZeroOrOne("w:startOverride", successors=("w:lvl",))
    ilvl = RequiredAttribute("w:ilvl", ST_DecimalNumber)

    def add_startOverride(self, val):
        """Return a newly added CT_DecimalNumber element having tagname
        ``w:startOverride`` and ``val`` attribute set to `val`."""
        return self._add_startOverride(val=val)


class CT_NumPr(BaseOxmlElement):
    """A ``<w:numPr>`` element, a container for numbering properties applied to a
    paragraph."""

    ilvl: CT_DecimalNumber | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:ilvl", successors=("w:numId", "w:numberingChange", "w:ins")
    )
    numId: CT_DecimalNumber | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:numId", successors=("w:numberingChange", "w:ins")
    )

    @property
    def ilvl_val(self) -> int | None:
        """Integer value of `w:ilvl/@w:val`, or |None| if not present."""
        ilvl = self.ilvl
        if ilvl is None:
            return None
        return ilvl.val

    @ilvl_val.setter
    def ilvl_val(self, value: int | None) -> None:
        if value is None:
            self._remove_ilvl()
            return
        ilvl = self.get_or_add_ilvl()
        ilvl.val = value

    @property
    def numId_val(self) -> int | None:
        """Integer value of `w:numId/@w:val`, or |None| if not present."""
        numId = self.numId
        if numId is None:
            return None
        return numId.val

    @numId_val.setter
    def numId_val(self, value: int | None) -> None:
        if value is None:
            self._remove_numId()
            return
        numId = self.get_or_add_numId()
        numId.val = value


class CT_Numbering(BaseOxmlElement):
    """``<w:numbering>`` element, the root element of a numbering part, i.e.
    numbering.xml."""

    abstractNum_lst: list[CT_AbstractNum]
    num_lst: list[CT_Num]
    numPicBullet_lst: list[CT_NumPicBullet]

    numPicBullet = ZeroOrMore(
        "w:numPicBullet", successors=("w:abstractNum", "w:numIdMacAtCleanup", "w:num")
    )
    abstractNum = ZeroOrMore(
        "w:abstractNum",
        successors=("w:numIdMacAtCleanup", "w:num"),
    )
    num = ZeroOrMore("w:num", successors=("w:numIdMacAtCleanup",))

    def add_num(self, abstractNum_id: int, num_id: int | None = None) -> CT_Num:
        """Return a newly added CT_Num (<w:num>) element referencing the abstract
        numbering definition identified by `abstractNum_id`.

        When `num_id` is not supplied, the next unused ``numId`` is chosen
        automatically.
        """
        if num_id is None:
            num_id = self._next_numId
        num = CT_Num.new(num_id, abstractNum_id)
        return self._insert_num(num)

    def add_abstractNum(self, abstractNum_id: int | None = None) -> CT_AbstractNum:
        """Return a newly added ``<w:abstractNum>`` child element.

        When `abstractNum_id` is not supplied, the next unused ``abstractNumId`` is
        chosen automatically.
        """
        if abstractNum_id is None:
            abstractNum_id = self._next_abstractNumId
        abstractNum = cast(
            CT_AbstractNum,
            parse_xml(
                f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{abstractNum_id}"/>'
            ),
        )
        # -- `abstractNum` elements must appear before any `num` elements --
        existing_abstracts = self.xpath("./w:abstractNum")
        if existing_abstracts:
            existing_abstracts[-1].addnext(abstractNum)
        else:
            # -- insert at start, before any w:num children --
            num_children = self.xpath("./w:num")
            if num_children:
                num_children[0].addprevious(abstractNum)
            else:
                self.append(abstractNum)
        return abstractNum

    def abstractNum_having_abstractNumId(self, abstractNumId: int) -> CT_AbstractNum:
        """Return the ``<w:abstractNum>`` child with matching `abstractNumId`."""
        try:
            return self.xpath(
                "./w:abstractNum[@w:abstractNumId=$abstractNumId]",
                abstractNumId=str(abstractNumId),
            )[0]
        except IndexError:
            raise KeyError(
                "no <w:abstractNum> element with abstractNumId %d" % abstractNumId
            )

    def num_having_numId(self, numId):
        """Return the ``<w:num>`` child element having ``numId`` attribute matching
        `numId`."""
        try:
            return self.xpath("./w:num[@w:numId=$numId]", numId=str(numId))[0]
        except IndexError:
            raise KeyError("no <w:num> element with numId %d" % numId)

    @property
    def _next_numId(self) -> int:
        """The first ``numId`` unused by a ``<w:num>`` element, starting at 1 and
        filling any gaps in numbering between existing ``<w:num>`` elements.

        O(n) with an O(1) fast path when existing ``numId`` values are contiguous
        starting at 1 (the common case).  Falls back to a full set scan only when
        there is an actual gap to fill, preserving gap-fill semantics.
        """
        numId_strs = self.xpath("./w:num/@w:numId")
        if not numId_strs:
            return 1
        num_ids = [int(s) for s in numId_strs]
        n = len(num_ids)
        max_id = max(num_ids)
        # -- Fast path: IDs form the contiguous set {1..n} with no gaps. --
        # -- Sum of 1..n is n*(n+1)//2; if that matches and max == n then
        # -- every int in that range is present and the next free id is n+1.
        if max_id == n and sum(num_ids) == n * (n + 1) // 2:
            return n + 1
        # -- Slow path: there's a gap; find it with a set membership test. --
        num_id_set = set(num_ids)
        for candidate in range(1, max_id + 2):
            if candidate not in num_id_set:
                return candidate
        # -- Unreachable, but satisfies type checker. --
        return max_id + 1

    @property
    def _next_abstractNumId(self) -> int:
        """Return the first unused ``abstractNumId``, starting at 0.

        Fills any gap in existing numbering.  O(n) with an O(1) fast path when
        the existing ids form the contiguous set ``{0..n-1}``.
        """
        abstractNumId_strs = self.xpath("./w:abstractNum/@w:abstractNumId")
        if not abstractNumId_strs:
            return 0
        ids = [int(s) for s in abstractNumId_strs]
        n = len(ids)
        max_id = max(ids)
        # -- Fast path: contiguous {0..n-1} -> next is n. --
        if max_id == n - 1 and sum(ids) == n * (n - 1) // 2:
            return n
        # -- Slow path: find the gap. --
        id_set = set(ids)
        for candidate in range(0, max_id + 2):
            if candidate not in id_set:
                return candidate
        return max_id + 1
