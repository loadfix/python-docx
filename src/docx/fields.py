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
from docx.oxml.parser import OxmlElement

if TYPE_CHECKING:
    from docx.document import Document
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

    # -- cross-reference resolution ---------------------------------------

    def resolve(self, document: "Document") -> str:
        """Return best-effort resolved text for this field.

        For ``REF`` fields referencing a bookmark, the text between the
        matching ``w:bookmarkStart`` and ``w:bookmarkEnd`` is returned as a
        single concatenated string. Heading references (``REF _Ref12345``)
        work the same way — the target is still a bookmark, typically placed
        around the heading's run text.

        For ``PAGEREF`` fields, python-docx cannot compute real page numbers
        because there is no layout engine; this method returns the cached
        :attr:`result_text` when present, otherwise ``"?"``.

        For any other field type (including ``PAGE``, ``DATE``, ``SEQ``,
        etc.), the existing :attr:`result_text` is returned unchanged. This
        method never raises for unresolvable references; callers can detect
        "couldn't resolve" by comparing against the field's original
        :attr:`result_text`.
        """
        field_type = self.type
        if field_type == "PAGEREF":
            cached = self.result_text
            return cached if cached else "?"
        if field_type != "REF":
            return self.result_text

        bookmark_name = _parse_ref_bookmark_name(self.instruction)
        if not bookmark_name:
            return self.result_text

        text = _bookmark_text(document, bookmark_name)
        if text is None:
            return self.result_text
        return text

    def update_result_text(self, new_text: str) -> None:
        """Replace this field's rendered result with `new_text`.

        For a simple field (``w:fldSimple``), this removes any existing runs
        and inserts a single ``w:r/w:t`` child. For a complex field, all runs
        between the ``separate`` and ``end`` markers are removed and replaced
        with a single new run containing `new_text`, inserted immediately
        before the ``end`` marker's run.
        """
        if self._kind == "simple":
            self._update_simple_result(new_text)
        else:
            self._update_complex_result(new_text)

    def _update_simple_result(self, new_text: str) -> None:
        fldSimple = self._element
        # -- remove all existing runs --
        for r in fldSimple.xpath("./w:r"):
            fldSimple.remove(r)
        # -- append a single new run carrying the text --
        new_r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = new_text
        if new_text != new_text.strip():
            t.set(qn("xml:space"), "preserve")
        new_r.append(t)
        fldSimple.append(new_r)

    def _update_complex_result(self, new_text: str) -> None:
        """Replace runs between ``separate`` and ``end`` with a single run.

        If no ``separate`` marker exists the field has no result region; this
        is a no-op — there's nowhere to write the rendered text.
        """
        separate_run = None
        end_run = None
        for r in self._iter_runs_after_begin():
            for child in r:
                if child.tag != qn("w:fldChar"):
                    continue
                fld_type = child.get(qn("w:fldCharType"))
                if fld_type == "separate":
                    separate_run = r
                elif fld_type == "end":
                    end_run = r
                    break
            if end_run is not None:
                break

        if separate_run is None or end_run is None:
            return

        # -- remove every sibling run strictly between separate_run and end_run --
        parent = separate_run.getparent()
        if parent is None:
            return
        sep_index = list(parent).index(separate_run)
        end_index = list(parent).index(end_run)
        # -- remove back-to-front to keep index valid --
        for i in range(end_index - 1, sep_index, -1):
            child = parent[i]
            if child.tag == qn("w:r"):
                parent.remove(child)

        # -- insert a single new run carrying the rendered text before end_run --
        new_r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = new_text
        if new_text != new_text.strip():
            t.set(qn("xml:space"), "preserve")
        new_r.append(t)
        end_run.addprevious(new_r)


# -- module-level helpers -------------------------------------------------


def _parse_ref_bookmark_name(instruction: str) -> str | None:
    """Return the bookmark name from a ``REF`` / ``PAGEREF`` instruction.

    Tokens are split on whitespace. The first token is the field type and is
    skipped. Subsequent tokens starting with a backslash (switches like ``\\h``,
    ``\\p``, ``\\* MERGEFORMAT``) are ignored. The first non-switch,
    non-quoted-empty token is treated as the bookmark name. Returns ``None``
    when no such token is present.
    """
    tokens = instruction.split()
    if len(tokens) < 2:
        return None
    # -- skip type token; then skip switches and their arguments --
    i = 1
    skip_next = False
    while i < len(tokens):
        token = tokens[i]
        if skip_next:
            skip_next = False
            i += 1
            continue
        if token.startswith("\\"):
            # -- switches like `\* MERGEFORMAT` consume an argument; `\h`
            #    and `\p` don't. We conservatively skip only formatting
            #    switches that are known to take an argument. --
            if token in ("\\*", "\\@", "\\#", "\\f"):
                skip_next = True
            i += 1
            continue
        # -- strip surrounding quotes if present --
        return token.strip('"')
    return None


def _bookmark_text(document: "Document", name: str) -> str | None:
    """Return the concatenated text between the bookmark's start and end.

    Walks the body XML from the matching ``w:bookmarkStart`` element to the
    ``w:bookmarkEnd`` with the same id, collecting every ``w:t`` descendant's
    text along the way. Returns ``None`` when no bookmark with `name` exists.

    The walk tolerates bookmarks that span paragraphs or sit inside hyperlinks
    / fields because it iterates the flattened pre-order descendant sequence
    of the body.
    """
    bookmark = document.bookmarks.get(name)
    if bookmark is None:
        return None

    bookmark_id = str(bookmark.bookmark_id)
    body = document._element.body  # pyright: ignore[reportPrivateUsage]
    start = body.xpath(f".//w:bookmarkStart[@w:id='{bookmark_id}']")
    end = body.xpath(f".//w:bookmarkEnd[@w:id='{bookmark_id}']")
    if not start:
        return None
    start_elm = start[0]
    end_elm = end[0] if end else None

    # -- iterate every descendant of body in document order; collect w:t text
    #    that sits between the start and (if present) end markers. --
    collecting = False
    parts: list[str] = []
    for elm in body.iter():
        if elm is start_elm:
            collecting = True
            continue
        if end_elm is not None and elm is end_elm:
            break
        if not collecting:
            continue
        if elm.tag == qn("w:t"):
            parts.append(elm.text or "")
    return "".join(parts)
