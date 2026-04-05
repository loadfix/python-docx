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
    """``<w:lvl>`` element, defining the format of a single level in an abstract
    numbering definition."""

    start = ZeroOrOne("w:start", successors=("w:numFmt", "w:lvlRestart", "w:pStyle",
                                              "w:isLgl", "w:suff", "w:lvlText",
                                              "w:lvlPicBulletId", "w:legacy", "w:lvlJc",
                                              "w:pPr", "w:rPr"))
    numFmt = ZeroOrOne("w:numFmt", successors=("w:lvlRestart", "w:pStyle", "w:isLgl",
                                                "w:suff", "w:lvlText",
                                                "w:lvlPicBulletId", "w:legacy",
                                                "w:lvlJc", "w:pPr", "w:rPr"))
    lvlText = ZeroOrOne("w:lvlText", successors=("w:lvlPicBulletId", "w:legacy",
                                                  "w:lvlJc", "w:pPr", "w:rPr"))
    pPr = ZeroOrOne("w:pPr", successors=("w:rPr",))
    rPr = ZeroOrOne("w:rPr", successors=())

    ilvl: int = RequiredAttribute("w:ilvl", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]

    @property
    def numFmt_val(self) -> str | None:
        """The value of the ``w:numFmt/@w:val`` attribute, or |None|."""
        numFmt = self.numFmt
        if numFmt is None:
            return None
        return numFmt.attrib.get(qn("w:val"))

    @numFmt_val.setter
    def numFmt_val(self, value: str | None):
        if value is None:
            self._remove_numFmt()
            return
        numFmt = self.get_or_add_numFmt()
        numFmt.attrib[qn("w:val")] = value

    @property
    def lvlText_val(self) -> str | None:
        """The value of the ``w:lvlText/@w:val`` attribute, or |None|."""
        lvlText = self.lvlText
        if lvlText is None:
            return None
        return lvlText.attrib.get(qn("w:val"))

    @lvlText_val.setter
    def lvlText_val(self, value: str | None):
        if value is None:
            self._remove_lvlText()
            return
        lvlText = self.get_or_add_lvlText()
        lvlText.attrib[qn("w:val")] = value

    @property
    def start_val(self) -> int | None:
        """The value of the ``w:start/@w:val`` attribute, or |None|."""
        start = self.start
        if start is None:
            return None
        val_str = start.attrib.get(qn("w:val"))
        return int(val_str) if val_str is not None else None

    @start_val.setter
    def start_val(self, value: int | None):
        if value is None:
            self._remove_start()
            return
        start = self.get_or_add_start()
        start.attrib[qn("w:val")] = str(value)


class CT_AbstractNum(BaseOxmlElement):
    """``<w:abstractNum>`` element, defining an abstract numbering definition that
    specifies the appearance and behavior of a numbered list."""

    lvl_lst: List[CT_Lvl]

    lvl = ZeroOrMore("w:lvl", successors=())

    abstractNumId: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:abstractNumId", ST_DecimalNumber
    )

    def add_lvl(self, ilvl: int) -> CT_Lvl:
        """Return a newly added ``<w:lvl>`` element with ``ilvl`` attribute set to
        `ilvl`."""
        lvl = cast(CT_Lvl, OxmlElement("w:lvl"))
        lvl.attrib[qn("w:ilvl")] = str(ilvl)
        self.append(lvl)
        return lvl

    @classmethod
    def new(cls, abstractNumId: int) -> CT_AbstractNum:
        """Return a new ``<w:abstractNum>`` element with the given `abstractNumId`."""
        abstractNum = cast(
            CT_AbstractNum,
            OxmlElement("w:abstractNum", attrs={qn("w:abstractNumId"): str(abstractNumId)}),
        )
        return abstractNum


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

    @property
    def abstractNumId_val(self) -> int:
        """The value of the ``w:abstractNumId`` child element's ``val`` attribute."""
        return self.abstractNumId.val

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

    ilvl = ZeroOrOne("w:ilvl", successors=("w:numId", "w:numberingChange", "w:ins"))
    numId = ZeroOrOne("w:numId", successors=("w:numberingChange", "w:ins"))

    @property
    def ilvl_val(self) -> int | None:
        """The value of ``w:ilvl/@w:val`` or |None| if not present."""
        ilvl = self.ilvl
        if ilvl is None:
            return None
        return ilvl.val

    @ilvl_val.setter
    def ilvl_val(self, value: int | None):
        if value is None:
            self._remove_ilvl()
            return
        self.get_or_add_ilvl().val = value

    @property
    def numId_val(self) -> int | None:
        """The value of ``w:numId/@w:val`` or |None| if not present."""
        numId = self.numId
        if numId is None:
            return None
        return numId.val

    @numId_val.setter
    def numId_val(self, value: int | None):
        if value is None:
            self._remove_numId()
            return
        self.get_or_add_numId().val = value


class CT_Numbering(BaseOxmlElement):
    """``<w:numbering>`` element, the root element of a numbering part, i.e.
    numbering.xml."""

    abstractNum_lst: List[CT_AbstractNum]

    abstractNum = ZeroOrMore("w:abstractNum", successors=("w:num", "w:numIdMacAtCleanup"))
    num = ZeroOrMore("w:num", successors=("w:numIdMacAtCleanup",))

    def add_abstractNum(self, abstractNumId: int | None = None) -> CT_AbstractNum:
        """Return a newly added ``<w:abstractNum>`` element.

        If `abstractNumId` is not provided, the next available ID is used.
        """
        if abstractNumId is None:
            abstractNumId = self._next_abstractNumId
        abstractNum = CT_AbstractNum.new(abstractNumId)
        return self._insert_abstractNum(abstractNum)

    def add_num(self, abstractNum_id):
        """Return a newly added CT_Num (<w:num>) element referencing the abstract
        numbering definition identified by `abstractNum_id`."""
        next_num_id = self._next_numId
        num = CT_Num.new(next_num_id, abstractNum_id)
        return self._insert_num(num)

    def num_having_numId(self, numId):
        """Return the ``<w:num>`` child element having ``numId`` attribute matching
        `numId`."""
        xpath = './w:num[@w:numId="%d"]' % numId
        try:
            return self.xpath(xpath)[0]
        except IndexError:
            raise KeyError("no <w:num> element with numId %d" % numId)

    @property
    def _next_abstractNumId(self) -> int:
        """The first ``abstractNumId`` unused by an ``<w:abstractNum>`` element."""
        abstractNumId_strs = self.xpath("./w:abstractNum/@w:abstractNumId")
        ids = [int(s) for s in abstractNumId_strs]
        for n in range(len(ids) + 1):
            if n not in ids:
                return n
        return 0  # unreachable but satisfies type checker

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
