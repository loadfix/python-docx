"""Custom element classes related to the numbering part."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List, cast

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

    lvl_lst: List[CT_Lvl]
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

    abstractNum_lst: List[CT_AbstractNum]
    num_lst: List[CT_Num]

    abstractNum = ZeroOrMore(
        "w:abstractNum",
        successors=("w:numIdMacAtCleanup", "w:numPicBullet", "w:num"),
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
        filling any gaps in numbering between existing ``<w:num>`` elements."""
        numId_strs = self.xpath("./w:num/@w:numId")
        num_ids = [int(numId_str) for numId_str in numId_strs]
        for num in range(1, len(num_ids) + 2):
            if num not in num_ids:
                break
        return num

    @property
    def _next_abstractNumId(self) -> int:
        """Return the first unused ``abstractNumId``, starting at 0.

        Fills any gap in existing numbering.
        """
        abstractNumId_strs = self.xpath("./w:abstractNum/@w:abstractNumId")
        ids = [int(s) for s in abstractNumId_strs]
        for candidate in range(0, len(ids) + 1):
            if candidate not in ids:
                return candidate
        return 0
