"""Custom element classes related to the numbering part."""

from __future__ import annotations

from typing import Callable, List, cast

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.shared import CT_DecimalNumber
from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


class CT_Lvl(BaseOxmlElement):
    """``<w:lvl>`` element, defining the format for a single level in a numbering
    definition."""

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
    numFmt = ZeroOrOne("w:numFmt", successors=_tag_seq[2:])
    lvlText = ZeroOrOne("w:lvlText", successors=_tag_seq[7:])
    lvlJc = ZeroOrOne("w:lvlJc", successors=_tag_seq[10:])
    pPr = ZeroOrOne("w:pPr", successors=_tag_seq[11:])
    rPr = ZeroOrOne("w:rPr", successors=())

    get_or_add_start: Callable[[], CT_DecimalNumber]
    get_or_add_numFmt: Callable[[], BaseOxmlElement]
    get_or_add_lvlText: Callable[[], BaseOxmlElement]
    get_or_add_lvlJc: Callable[[], BaseOxmlElement]
    get_or_add_pPr: Callable[[], BaseOxmlElement]
    get_or_add_rPr: Callable[[], BaseOxmlElement]
    _remove_numFmt: Callable[[], None]
    _remove_lvlText: Callable[[], None]
    _remove_lvlJc: Callable[[], None]

    ilvl: int = RequiredAttribute("w:ilvl", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]

    del _tag_seq

    @property
    def start_val(self) -> int:
        """The value of ``<w:start>`` child ``val`` attribute, or 1 if not present."""
        start = self.start
        if start is None:
            return 1
        return start.val

    @start_val.setter
    def start_val(self, value: int) -> None:
        self.get_or_add_start().val = value

    @property
    def numFmt_val(self) -> str | None:
        """Value of ``<w:numFmt>`` child ``val`` attribute, or None."""
        numFmt = self.numFmt
        if numFmt is None:
            return None
        return numFmt.get(qn("w:val"))

    @numFmt_val.setter
    def numFmt_val(self, value: str | None) -> None:
        if value is None:
            self._remove_numFmt()
            return
        numFmt = self.get_or_add_numFmt()
        numFmt.set(qn("w:val"), value)

    @property
    def lvlText_val(self) -> str | None:
        """Value of ``<w:lvlText>`` child ``val`` attribute, or None."""
        lvlText = self.lvlText
        if lvlText is None:
            return None
        return lvlText.get(qn("w:val"))

    @lvlText_val.setter
    def lvlText_val(self, value: str | None) -> None:
        if value is None:
            self._remove_lvlText()
            return
        lvlText = self.get_or_add_lvlText()
        lvlText.set(qn("w:val"), value)


class CT_AbstractNum(BaseOxmlElement):
    """``<w:abstractNum>`` element, defining a numbering definition's formatting."""

    lvl = ZeroOrMore("w:lvl", successors=())

    lvl_lst: List[CT_Lvl]

    abstractNumId: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:abstractNumId", ST_DecimalNumber
    )

    @classmethod
    def new(cls, abstractNum_id: int) -> CT_AbstractNum:
        """Return a new ``<w:abstractNum>`` element with the given abstractNumId."""
        abstractNum = cast(
            CT_AbstractNum,
            OxmlElement("w:abstractNum", attrs={qn("w:abstractNumId"): str(abstractNum_id)}),
        )
        return abstractNum

    def add_lvl(self, ilvl: int) -> CT_Lvl:
        """Add a ``<w:lvl>`` child element with ``ilvl`` attribute set to `ilvl`."""
        lvl = cast(CT_Lvl, OxmlElement("w:lvl", attrs={qn("w:ilvl"): str(ilvl)}))
        self.append(lvl)
        return lvl

    def lvl_for_ilvl(self, ilvl: int) -> CT_Lvl | None:
        """Return the ``<w:lvl>`` element with matching `ilvl`, or None."""
        xpath = './w:lvl[@w:ilvl="%d"]' % ilvl
        results = self.xpath(xpath)
        return results[0] if results else None


class CT_Num(BaseOxmlElement):
    """``<w:num>`` element, which represents a concrete list definition instance, having
    a required child <w:abstractNumId> that references an abstract numbering definition
    that defines most of the formatting details."""

    abstractNumId = OneAndOnlyOne("w:abstractNumId")
    lvlOverride = ZeroOrMore("w:lvlOverride")
    numId = RequiredAttribute("w:numId", ST_DecimalNumber)

    lvlOverride_lst: List[CT_NumLvl]

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

    @property
    def abstractNumId_val(self) -> int:
        """The value of the ``<w:abstractNumId>`` child ``val`` attribute."""
        return self.abstractNumId.val


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

    get_or_add_ilvl: Callable[[], CT_DecimalNumber]
    get_or_add_numId: Callable[[], CT_DecimalNumber]
    _remove_ilvl: Callable[[], None]
    _remove_numId: Callable[[], None]

    ilvl: CT_DecimalNumber | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:ilvl", successors=("w:numId", "w:numberingChange", "w:ins")
    )
    numId: CT_DecimalNumber | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:numId", successors=("w:numberingChange", "w:ins")
    )

    @property
    def ilvl_val(self) -> int | None:
        """Value of ``<w:ilvl>`` child ``val`` attribute, or None if not present."""
        ilvl = self.ilvl
        if ilvl is None:
            return None
        return ilvl.val

    @ilvl_val.setter
    def ilvl_val(self, value: int | None) -> None:
        if value is None:
            self._remove_ilvl()
            return
        self.get_or_add_ilvl().val = value

    @property
    def numId_val(self) -> int | None:
        """Value of ``<w:numId>`` child ``val`` attribute, or None if not present."""
        numId = self.numId
        if numId is None:
            return None
        return numId.val

    @numId_val.setter
    def numId_val(self, value: int | None) -> None:
        if value is None:
            self._remove_numId()
            return
        self.get_or_add_numId().val = value


class CT_Numbering(BaseOxmlElement):
    """``<w:numbering>`` element, the root element of a numbering part, i.e.
    numbering.xml."""

    abstractNum = ZeroOrMore("w:abstractNum", successors=("w:num", "w:numIdMacAtCleanup"))
    num = ZeroOrMore("w:num", successors=("w:numIdMacAtCleanup",))

    abstractNum_lst: List[CT_AbstractNum]
    num_lst: List[CT_Num]

    def add_num(self, abstractNum_id):
        """Return a newly added CT_Num (<w:num>) element referencing the abstract
        numbering definition identified by `abstractNum_id`."""
        next_num_id = self._next_numId
        num = CT_Num.new(next_num_id, abstractNum_id)
        return self._insert_num(num)

    def add_abstractNum(self, abstractNum: CT_AbstractNum) -> CT_AbstractNum:
        """Append ``abstractNum`` element, inserting in proper sequence."""
        return self._insert_abstractNum(abstractNum)

    def num_having_numId(self, numId):
        """Return the ``<w:num>`` child element having ``numId`` attribute matching
        `numId`."""
        xpath = './w:num[@w:numId="%d"]' % numId
        try:
            return self.xpath(xpath)[0]
        except IndexError:
            raise KeyError("no <w:num> element with numId %d" % numId)

    def abstractNum_having_abstractNumId(self, abstractNumId: int) -> CT_AbstractNum:
        """Return the ``<w:abstractNum>`` child with matching ``abstractNumId``."""
        xpath = './w:abstractNum[@w:abstractNumId="%d"]' % abstractNumId
        try:
            return self.xpath(xpath)[0]
        except IndexError:
            raise KeyError("no <w:abstractNum> element with abstractNumId %d" % abstractNumId)

    @property
    def _next_numId(self):
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
        """The first ``abstractNumId`` unused by a ``<w:abstractNum>`` element."""
        id_strs = self.xpath("./w:abstractNum/@w:abstractNumId")
        ids = [int(id_str) for id_str in id_strs]
        for n in range(0, len(ids) + 1):
            if n not in ids:
                return n
