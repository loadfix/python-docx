"""Custom element classes related to field codes.

Word supports two XML forms for field codes:

* "Simple" fields use a single `<w:fldSimple>` element containing the rendered
  result as one or more `<w:r>` children; the instruction string is stored in
  the `w:instr` attribute.
* "Complex" fields span multiple runs using `<w:fldChar>` markers (``begin``,
  ``separate``, ``end``) and an `<w:instrText>` run containing the instruction
  string.

Both forms surface through the same :class:`docx.fields.Field` proxy.
"""

from __future__ import annotations

from typing import TYPE_CHECKING
from collections.abc import Callable

from docx.oxml.simpletypes import XsdString
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    RequiredAttribute,
    ZeroOrMore,
)

if TYPE_CHECKING:
    from docx.oxml.text.run import CT_R


class ST_FldCharType(XsdString):
    """Valid values for the `w:fldCharType` attribute."""

    @classmethod
    def validate(cls, value: str) -> None:
        cls.validate_string(value)
        valid_values = ("begin", "separate", "end")
        if value not in valid_values:
            raise ValueError(
                "w:fldCharType must be one of %s, got '%s'" % (valid_values, value)
            )


class CT_FldSimple(BaseOxmlElement):
    """`<w:fldSimple>` element, a "simple" (one-element) field code.

    It is a block-level child of `w:p`, containing one or more `<w:r>` children
    that hold the current rendered result. The field instruction is stored in
    the `w:instr` attribute.
    """

    add_r: Callable[[], "CT_R"]
    r_lst: list["CT_R"]

    instr: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:instr", XsdString
    )

    r = ZeroOrMore("w:r")

    @property
    def text(self) -> str:  # pyright: ignore[reportIncompatibleMethodOverride]
        """The textual content of this simple field (the rendered result).

        Concatenates the text of all `w:r` children as well as any nested
        `w:fldSimple` descendants (nested fields are uncommon but permitted).
        """
        return "".join(r.text for r in self.xpath(".//w:r"))


class CT_FldChar(BaseOxmlElement):
    """`<w:fldChar>` element, a complex field begin/separate/end marker.

    Occurs as a child of `<w:r>`. The `w:fldCharType` attribute indicates which
    of the three roles it plays.
    """

    fldCharType: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:fldCharType", ST_FldCharType
    )


class CT_InstrText(BaseOxmlElement):
    """`<w:instrText>` element, containing the field instruction for a complex
    field. Occurs as a child of `<w:r>`.
    """

    def __str__(self) -> str:
        """Text contained in this element, the empty string if it has no content."""
        return self.text or ""
