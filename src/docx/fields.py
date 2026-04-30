"""Field-related proxy objects and field-type constants.

A "field" in WordprocessingML is an instruction (e.g. ``PAGE`` or
``REF bookmark1 \\h``) that Word evaluates at display time to produce some
rendered text. Two XML forms are supported:

* **Simple fields** — a single ``<w:fldSimple>`` element whose `w:instr`
  attribute holds the instruction and whose child runs hold the most-recently
  rendered result.
* **Complex fields** — a sequence of ``<w:r>`` runs delimited by
  ``<w:fldChar>`` markers (``begin``, ``separate``, ``end``) with the
  instruction stored in an ``<w:instrText>`` element between ``begin`` and
  ``separate``, and the rendered result as ordinary text between ``separate``
  and ``end``.

Both forms surface through the same :class:`Field` proxy.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.oxml.ns import qn

if TYPE_CHECKING:
    from docx.oxml.fields import CT_FldSimple
    from docx.oxml.text.run import CT_R
    from docx.oxml.xmlchemy import BaseOxmlElement


class WD_FIELD_TYPE:
    """Common field-type identifiers (the first token of a field instruction).

    Usage::

        paragraph.add_simple_field(f"{WD_FIELD_TYPE.PAGE}", "1")

    These are plain string constants rather than an `enum.Enum` because field
    types are open-ended — callers can use any string (e.g. a custom field) and
    readers will correctly populate :attr:`Field.type` from whatever is found in
    the document. The enum-ish class is for autocompletion and typo avoidance.
    """

    PAGE = "PAGE"
    NUMPAGES = "NUMPAGES"
    DATE = "DATE"
    TIME = "TIME"
    AUTHOR = "AUTHOR"
    REF = "REF"
    TOC = "TOC"
    SEQ = "SEQ"
    HYPERLINK = "HYPERLINK"
    PAGEREF = "PAGEREF"


class Field:
    """Proxy for a field in a paragraph.

    A :class:`Field` wraps either a ``<w:fldSimple>`` element (simple form) or
    the opening ``<w:r>`` run containing the ``begin`` ``<w:fldChar>`` marker
    (complex form). Both forms expose the same three read-only properties:

    * :attr:`instruction` — the raw instruction text
    * :attr:`type` — the first whitespace-delimited token of the instruction
    * :attr:`result_text` — the most recently computed rendered result, or the
      empty string when absent
    """

    def __init__(self, kind: str, element: "BaseOxmlElement"):
        self._kind = kind
        self._element = element

    @classmethod
    def for_simple(cls, fldSimple: "CT_FldSimple") -> "Field":
        """Return a :class:`Field` wrapping a ``w:fldSimple`` element."""
        return cls("simple", fldSimple)

    @classmethod
    def for_complex(cls, begin_run: "CT_R") -> "Field":
        """Return a :class:`Field` wrapping the ``begin`` run of a complex field."""
        return cls("complex", begin_run)

    @property
    def is_complex(self) -> bool:
        """``True`` for a complex (three-marker) field, ``False`` for simple."""
        return self._kind == "complex"

    @property
    def instruction(self) -> str:
        """The raw instruction text of this field.

        For simple fields this is the `w:instr` attribute value. For complex
        fields this is the concatenated text of all ``<w:instrText>`` runs
        between the ``begin`` and ``separate`` markers (or end-of-paragraph if
        no ``separate`` marker is present).
        """
        if self._kind == "simple":
            return self._element.get(qn("w:instr")) or ""
        return self._read_complex_instruction()

    @property
    def type(self) -> str:
        """The first whitespace-delimited token of :attr:`instruction`.

        For ``"REF bookmark1 \\h"`` this returns ``"REF"``. The empty string is
        returned when the instruction is empty or whitespace-only.
        """
        instr = self.instruction.strip()
        if not instr:
            return ""
        return instr.split()[0]

    @property
    def result_text(self) -> str:
        """The rendered result text for this field.

        For simple fields this is the text of any runs nested in the
        ``<w:fldSimple>`` element. For complex fields this is the text of runs
        between the ``separate`` and ``end`` markers. The empty string is
        returned when no result is available (for example a complex field with
        no ``separate`` marker).
        """
        if self._kind == "simple":
            return self._read_simple_result()
        return self._read_complex_result()

    # -- internals ---------------------------------------------------------

    def _read_simple_result(self) -> str:
        """Concatenate the text of every ``w:r`` descendant of the fldSimple."""
        return "".join(r.text for r in self._element.xpath(".//w:r"))

    def _read_complex_instruction(self) -> str:
        """Walk runs after the begin marker, concatenating ``w:instrText`` until
        the first ``separate`` or ``end`` marker."""
        parts: list[str] = []
        for r in self._iter_runs_after_begin():
            for child in r:
                tag = child.tag
                if tag == qn("w:fldChar"):
                    fld_type = child.get(qn("w:fldCharType"))
                    if fld_type in ("separate", "end"):
                        return "".join(parts)
                elif tag == qn("w:instrText"):
                    parts.append(child.text or "")
        return "".join(parts)

    def _read_complex_result(self) -> str:
        """Walk runs after the begin marker, finding ``separate``, then
        concatenating run text until the first ``end`` marker."""
        seen_separate = False
        parts: list[str] = []
        for r in self._iter_runs_after_begin():
            for child in r:
                tag = child.tag
                if tag == qn("w:fldChar"):
                    fld_type = child.get(qn("w:fldCharType"))
                    if fld_type == "separate":
                        seen_separate = True
                        break
                    if fld_type == "end":
                        return "".join(parts)
            else:
                # -- no fldChar encountered in this run --
                if seen_separate:
                    parts.append(r.text or "")
        return "".join(parts)

    def _iter_runs_after_begin(self):
        """Yield each ``w:r`` sibling following the begin-run in document order."""
        sibling = self._element.getnext()
        while sibling is not None:
            if sibling.tag == qn("w:r"):
                yield sibling
            sibling = sibling.getnext()
