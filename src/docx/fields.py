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

import ast
import re
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

    .. versionadded:: 2026.05.0
    """

    PAGE = "PAGE"
    NUMPAGES = "NUMPAGES"
    DATE = "DATE"
    TIME = "TIME"
    AUTHOR = "AUTHOR"
    TITLE = "TITLE"
    FILENAME = "FILENAME"
    REF = "REF"
    TOC = "TOC"
    SEQ = "SEQ"
    HYPERLINK = "HYPERLINK"
    PAGEREF = "PAGEREF"
    NOTEREF = "NOTEREF"
    SEQREF = "SEQREF"
    MERGEFIELD = "MERGEFIELD"
    STYLEREF = "STYLEREF"
    NUMBEREDHEADERS = "NUMBEREDHEADERS"


class Field:
    """Proxy for a field in a paragraph.

    A :class:`Field` wraps either a ``<w:fldSimple>`` element (simple form) or
    the opening ``<w:r>`` run containing the ``begin`` ``<w:fldChar>`` marker
    (complex form). Both forms expose the same three read-only properties:

    * :attr:`instruction` — the raw instruction text
    * :attr:`type` — the first whitespace-delimited token of the instruction
    * :attr:`result_text` — the most recently computed rendered result, or the
      empty string when absent

    .. versionadded:: 2026.05.0
    """

    def __init__(self, kind: str, element: "BaseOxmlElement"):
        self._kind = kind
        self._element = element

    @classmethod
    def for_simple(cls, fldSimple: "CT_FldSimple") -> "Field":
        """Return a :class:`Field` wrapping a ``w:fldSimple`` element.

        .. versionadded:: 2026.05.0
        """
        return cls("simple", fldSimple)

    @classmethod
    def for_complex(cls, begin_run: "CT_R") -> "Field":
        """Return a :class:`Field` wrapping the ``begin`` run of a complex field.

        .. versionadded:: 2026.05.0
        """
        return cls("complex", begin_run)

    @property
    def is_complex(self) -> bool:
        """``True`` for a complex (three-marker) field, ``False`` for simple.

        .. versionadded:: 2026.05.0
        """
        return self._kind == "complex"

    @property
    def instruction(self) -> str:
        """The raw instruction text of this field.

        For simple fields this is the `w:instr` attribute value. For complex
        fields this is the concatenated text of all ``<w:instrText>`` runs
        between the ``begin`` and ``separate`` markers (or end-of-paragraph if
        no ``separate`` marker is present).

        .. versionadded:: 2026.05.0
        """
        if self._kind == "simple":
            return self._element.get(qn("w:instr")) or ""
        return self._read_complex_instruction()

    @property
    def type(self) -> str:
        """The first whitespace-delimited token of :attr:`instruction`.

        For ``"REF bookmark1 \\h"`` this returns ``"REF"``. The empty string is
        returned when the instruction is empty or whitespace-only.

        .. versionadded:: 2026.05.0
        """
        instr = self.instruction.strip()
        if not instr:
            return ""
        return instr.split()[0]

    @property
    def field_type(self) -> str:
        """Alias for :attr:`type`. The first token of the instruction.

        Provided for readability — ``field.type`` is ambiguous in code that also
        deals with Python types. Both spellings return the same value.

        .. versionadded:: 2026.05.10
        """
        return self.type

    @property
    def result_text(self) -> str:
        """The rendered result text for this field.

        For simple fields this is the text of any runs nested in the
        ``<w:fldSimple>`` element. For complex fields this is the text of runs
        between the ``separate`` and ``end`` markers. The empty string is
        returned when no result is available (for example a complex field with
        no ``separate`` marker).

        .. versionadded:: 2026.05.0
        """
        if self._kind == "simple":
            return self._read_simple_result()
        return self._read_complex_result()

    @property
    def result(self) -> str:
        """Alias for :attr:`result_text`. The rendered result of the field.

        .. versionadded:: 2026.05.10
        """
        return self.result_text

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

        For property-lookup fields — ``DOCPROPERTY``, ``AUTHOR``, ``TITLE``,
        ``SUBJECT``, ``KEYWORDS``, ``COMMENTS``, ``LASTSAVEDBY`` — the value
        is resolved against the document's :class:`CoreProperties` (for the
        well-known names) and :class:`CustomProperties` (for ``DOCPROPERTY``
        with a custom name). Closes upstream#1482.

        For any other field type (including ``PAGE``, ``DATE``, ``SEQ``,
        etc.), the existing :attr:`result_text` is returned unchanged. This
        method never raises for unresolvable references; callers can detect
        "couldn't resolve" by comparing against the field's original
        :attr:`result_text`.

        .. versionadded:: 2026.05.0
        """
        field_type = self.type
        if field_type == "PAGEREF":
            cached = self.result_text
            return cached if cached else "?"
        if field_type in _PROPERTY_FIELD_TYPES:
            resolved = _resolve_property_field(self, document)
            if resolved is not None:
                return resolved
            return self.result_text
        if field_type != "REF":
            return self.result_text

        bookmark_name = _parse_ref_bookmark_name(self.instruction)
        if not bookmark_name:
            return self.result_text

        text = _bookmark_text(document, bookmark_name)
        if text is None:
            return self.result_text
        return text

    def evaluate(self, context: "dict[str, object] | None" = None) -> str:
        """Return the best-effort evaluated text for this field.

        Extends :meth:`resolve` to a wider subset of field types, driven by a
        caller-supplied ``context`` mapping (typically mail-merge values):

        * ``MERGEFIELD name`` — looked up in ``context`` by name; if the key
          is missing the cached :attr:`result_text` is returned unchanged.
        * ``IF expr1 op expr2 "true-text" "false-text"`` — ``expr1`` and
          ``expr2`` may be quoted string literals, nested ``{MERGEFIELD name}``
          references, or bare names resolved from ``context``. Supported
          operators are ``=``, ``<>``, ``!=``, ``<``, ``>``, ``<=``, ``>=``.
          Numeric comparison is used when both sides parse as numbers;
          otherwise a case-sensitive string compare is used.
        * ``HYPERLINK "url"`` — returns the URL argument (or the cached
          :attr:`result_text` when present and non-empty; that is the display
          text Word already rendered).
        * ``= <expr>`` formula — evaluated as a restricted arithmetic
          expression over ``+``, ``-``, ``*``, ``/``, ``%``, ``**`` and
          parentheses. Nested ``{MERGEFIELD name}`` references are substituted
          from ``context`` before evaluation. Returns the string form of the
          result, or the cached :attr:`result_text` on parse/eval error.
        * ``PAGE`` / ``NUMPAGES`` / ``DATE`` / ``TIME`` — runtime-dynamic
          fields that python-docx cannot compute. Returns the cached
          :attr:`result_text` when present, otherwise the sentinel ``"?"``.
        * ``REF`` / ``PAGEREF`` / ``DOCPROPERTY`` / core-property fields —
          delegated to :meth:`resolve` against the owning document discovered
          through the element tree (``context["document"]`` if provided).

        Any other field type returns the existing :attr:`result_text`
        unchanged. This method never raises for unresolvable references.

        .. versionadded:: 2026.05.7
        """
        ctx: dict[str, object] = dict(context) if context else {}
        field_type = self.type

        if field_type in _PROPERTY_FIELD_TYPES or field_type in ("REF", "PAGEREF"):
            document = ctx.get("document")
            if document is not None:
                try:
                    return self.resolve(document)  # type: ignore[arg-type]
                except Exception:
                    return self.result_text
            return self.result_text

        if field_type == "MERGEFIELD":
            name = _parse_mergefield_name(self.instruction)
            if name is None:
                return self.result_text
            value = ctx.get(name)
            return self.result_text if value is None else str(value)

        if field_type == "IF":
            evaluated = _evaluate_if(self.instruction, ctx)
            return self.result_text if evaluated is None else evaluated

        if field_type == "HYPERLINK":
            url = _parse_hyperlink_url(self.instruction)
            cached = self.result_text
            if cached:
                return cached
            return url if url is not None else self.result_text

        if field_type == "=":
            evaluated = _evaluate_formula(self.instruction, ctx)
            return self.result_text if evaluated is None else evaluated

        if field_type in _RUNTIME_DYNAMIC_FIELD_TYPES:
            cached = self.result_text
            return cached if cached else "?"

        return self.result_text

    def mark_dirty(self) -> None:
        """Mark this field's cached result as stale (``@w:dirty="true"``).

        Word consults the ``w:dirty`` attribute on the field's ``w:fldChar``
        begin marker (complex fields) or on the ``w:fldSimple`` element
        (simple fields) to decide whether to re-evaluate the field on open.
        Setting this flag is the programmatic equivalent of right-clicking
        the field in Word and choosing *Update Field*: Word will recompute
        the result the next time the document is opened or refreshed.

        This is especially useful for ``TOC`` fields where the cached
        preview python-docx produces is not a real rendering — marking the
        TOC dirty forces Word to rebuild it on open. Closes upstream#1403.

        .. versionadded:: 2026.05.0
        """
        if self._kind == "simple":
            self._element.set(qn("w:dirty"), "true")
            return
        # -- complex field: find the `w:fldChar` with @fldCharType="begin"
        #    inside the begin-run and set its `w:dirty` attribute --
        for child in self._element:
            if child.tag != qn("w:fldChar"):
                continue
            if child.get(qn("w:fldCharType")) == "begin":
                child.set(qn("w:dirty"), "true")
                return

    @property
    def as_cross_reference(self) -> "CrossReference | None":
        """Return a :class:`CrossReference` view of this field, or |None|.

        Returns |None| when :attr:`type` is not one of ``REF``, ``PAGEREF``,
        ``NOTEREF``, ``SEQREF``, ``STYLEREF``. The returned object shares
        the same underlying XML element, so writes (e.g.
        :meth:`update_result_text`) propagate back to the document.

        .. versionadded:: 2026.05.10
        """
        return _wrap_as_cross_reference(self)

    @property
    def as_toc(self) -> "TocField | None":
        """Return a :class:`TocField` view of this field, or |None|.

        Returns a :class:`TableOfFiguresField` when the instruction is a
        ``TOC \\c "label"`` variant, a :class:`TableOfAuthoritiesField`
        when the field type is ``TOA``, a plain :class:`TocField` for
        ordinary ``TOC`` fields, and |None| for everything else. The
        returned object shares the same underlying XML element, so writes
        (e.g. :meth:`update_result_text`, :meth:`mark_dirty`) propagate
        back to the document.

        .. versionadded:: 2026.05.10
        """
        return _wrap_as_toc(self)

    @property
    def is_dirty(self) -> bool:
        """True when this field is marked dirty (``@w:dirty="true"``).

        For simple fields this reads ``w:fldSimple/@w:dirty``; for complex
        fields it reads the ``@w:dirty`` attribute on the ``begin``
        ``w:fldChar`` marker. Values are interpreted using ST_OnOff
        semantics: ``"true"``, ``"1"``, and ``"on"`` (case-insensitive)
        count as true; anything else (including the attribute's absence)
        as false.

        .. versionadded:: 2026.05.0
        """
        if self._kind == "simple":
            val = self._element.get(qn("w:dirty"))
        else:
            val = None
            for child in self._element:
                if child.tag != qn("w:fldChar"):
                    continue
                if child.get(qn("w:fldCharType")) == "begin":
                    val = child.get(qn("w:dirty"))
                    break
        if val is None:
            return False
        return val.strip().lower() in ("true", "1", "on")

    def update_result_text(self, new_text: str) -> None:
        """Replace this field's rendered result with `new_text`.

        For a simple field (``w:fldSimple``), this removes any existing runs
        and inserts a single ``w:r/w:t`` child. For a complex field, all runs
        between the ``separate`` and ``end`` markers are removed and replaced
        with a single new run containing `new_text`, inserted immediately
        before the ``end`` marker's run.

        .. versionadded:: 2026.05.0
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


class ParsedFieldInstruction:
    """Structured view of a parsed field instruction.

    Produced by :func:`parse_field_instruction`. Carries the field-type name,
    the positional arguments (quoted strings unquoted, nested ``{...}``
    groups preserved verbatim), and the formatting switches — a mapping from
    switch token (without the leading backslash, uppercased) to its argument,
    or to the empty string when the switch takes no argument. Switches that
    take an argument per ECMA-376 § 17.16 are: ``\\*`` (general format),
    ``\\@`` (date/time picture), ``\\#`` (numeric picture), and ``\\f``
    (font). All other switches are flag-only.

    .. versionadded:: 2026.05.10
    """

    __slots__ = ("name", "args", "switches")

    def __init__(
        self,
        name: str,
        args: "list[str]",
        switches: "dict[str, str]",
    ):
        self.name = name
        self.args = args
        self.switches = switches

    def __eq__(self, other: object) -> bool:  # pragma: no cover - trivial
        if not isinstance(other, ParsedFieldInstruction):
            return NotImplemented
        return (
            self.name == other.name
            and self.args == other.args
            and self.switches == other.switches
        )

    def __repr__(self) -> str:  # pragma: no cover - trivial
        return (
            f"ParsedFieldInstruction(name={self.name!r}, "
            f"args={self.args!r}, switches={self.switches!r})"
        )


# -- switches that take an argument, per ECMA-376 § 17.16.4 --
_ARG_TAKING_SWITCHES = frozenset({"*", "@", "#", "f"})


def parse_field_instruction(instruction: str) -> ParsedFieldInstruction:
    """Parse a field instruction string into (name, args, switches).

    ``"MERGEFIELD FirstName \\* MERGEFORMAT"`` →
    ``ParsedFieldInstruction(name="MERGEFIELD", args=["FirstName"],
    switches={"*": "MERGEFORMAT"})``.

    The instruction is tokenised honouring double-quoted runs (``"a b c"``
    becomes a single ``a b c`` argument) and brace-delimited nested-field
    groups (``{MERGEFIELD foo}`` is kept as a single atomic token). Switches
    are distinguished by a leading backslash. Switches in the set
    ``\\*``, ``\\@``, ``\\#``, ``\\f`` consume the next token as their
    argument; all other switches are recorded as flag-only (``switches[key]``
    is the empty string).

    Returns an empty :class:`ParsedFieldInstruction` (``name=""``, no args,
    no switches) when the instruction is empty or whitespace-only. Does not
    raise for malformed input: unknown switch tokens are recorded verbatim.

    .. versionadded:: 2026.05.10
    """
    stripped = instruction.strip()
    if not stripped:
        return ParsedFieldInstruction(name="", args=[], switches={})

    tokens = _tokenize_field_args(stripped)
    if not tokens:
        return ParsedFieldInstruction(name="", args=[], switches={})

    name = tokens[0]
    args: list[str] = []
    switches: dict[str, str] = {}

    i = 1
    while i < len(tokens):
        token = tokens[i]
        if token.startswith("\\") and len(token) >= 2:
            # -- strip leading backslash and uppercase the letter switches so
            #    the common ones (`\*`, `\@`, `\#`, `\f`, `\h`, `\p`, `\l`)
            #    are case-insensitive on the value side. Single-char switches
            #    are left as-is (case-insensitive for ASCII letters). --
            raw_key = token[1:]
            key = raw_key.upper() if raw_key.isalpha() else raw_key
            # -- is this switch argument-taking? use the lowercase form so
            #    `\*` / `\@` / `\#` / `\f` match regardless of case. --
            tag = raw_key.lower() if raw_key.isalpha() else raw_key
            if tag in _ARG_TAKING_SWITCHES and i + 1 < len(tokens):
                switches[key] = tokens[i + 1]
                i += 2
                continue
            switches[key] = ""
            i += 1
            continue
        args.append(token)
        i += 1

    return ParsedFieldInstruction(name=name, args=args, switches=switches)


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


_PROPERTY_FIELD_TYPES = frozenset(
    {
        "DOCPROPERTY",
        "AUTHOR",
        "TITLE",
        "SUBJECT",
        "KEYWORDS",
        "COMMENTS",
        "LASTSAVEDBY",
    }
)

# -- mapping: DOCPROPERTY name (Word's built-in labels) -> CoreProperties attr --
_CORE_PROPERTY_ATTR_MAP = {
    "Author": "author",
    "Title": "title",
    "Subject": "subject",
    "Keywords": "keywords",
    "Comments": "comments",
    "Category": "category",
    "LastSavedBy": "last_modified_by",
    "ContentStatus": "content_status",
    "Language": "language",
    "Version": "version",
    "RevisionNumber": "revision",
}

# -- field-type-token -> CoreProperties attribute for bare-name fields --
_CORE_FIELD_TYPE_ATTR_MAP = {
    "AUTHOR": "author",
    "TITLE": "title",
    "SUBJECT": "subject",
    "KEYWORDS": "keywords",
    "COMMENTS": "comments",
    "LASTSAVEDBY": "last_modified_by",
}


def _parse_docproperty_name(instruction: str) -> str | None:
    """Return the property name argument from a ``DOCPROPERTY`` instruction.

    Walks the whitespace-split tokens after ``DOCPROPERTY``, skipping
    formatting switches and reassembling quoted multi-word names. Returns
    |None| when no name is found.
    """
    # -- extract the substring after "DOCPROPERTY" --
    stripped = instruction.strip()
    if not stripped.upper().startswith("DOCPROPERTY"):
        return None
    remainder = stripped[len("DOCPROPERTY") :].strip()
    if not remainder:
        return None

    # -- quoted form: "Some Name" [switches...] --
    if remainder.startswith('"'):
        end = remainder.find('"', 1)
        if end == -1:
            return None
        return remainder[1:end]

    # -- otherwise, split on whitespace; first non-switch token is the name --
    tokens = remainder.split()
    for token in tokens:
        if token.startswith("\\"):
            continue
        return token
    return None


def _resolve_property_field(field: "Field", document: "Document") -> str | None:
    """Return the resolved value for a property field, or |None| when unresolved."""
    field_type = field.type
    if field_type == "DOCPROPERTY":
        prop_name = _parse_docproperty_name(field.instruction)
        if not prop_name:
            return None
        # -- try CoreProperties aliases first --
        core_attr = _CORE_PROPERTY_ATTR_MAP.get(prop_name)
        if core_attr is not None:
            value = getattr(document.core_properties, core_attr, None)
            if value is not None:
                return str(value)
        # -- fall back to CustomProperties --
        try:
            custom = document.custom_properties
        except Exception:
            return None
        value = custom.get(prop_name)
        if value is None:
            return None
        return str(value)

    core_attr = _CORE_FIELD_TYPE_ATTR_MAP.get(field_type)
    if core_attr is None:
        return None
    value = getattr(document.core_properties, core_attr, None)
    return None if value is None else str(value)


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


# -- evaluation helpers ---------------------------------------------------


_RUNTIME_DYNAMIC_FIELD_TYPES = frozenset(
    {"PAGE", "NUMPAGES", "DATE", "TIME", "SECTIONPAGES", "SECTION"}
)


def _parse_mergefield_name(instruction: str) -> str | None:
    """Return the field-name argument from a ``MERGEFIELD`` instruction.

    Handles both bare (``MERGEFIELD foo``) and quoted (``MERGEFIELD "foo bar"``)
    forms. Formatting switches (``\\* MERGEFORMAT``, ``\\b``, ``\\f``) are
    skipped. Returns |None| when no name can be extracted.
    """
    stripped = instruction.strip()
    if not stripped.upper().startswith("MERGEFIELD"):
        return None
    remainder = stripped[len("MERGEFIELD") :].strip()
    if not remainder:
        return None
    if remainder.startswith('"'):
        end = remainder.find('"', 1)
        if end == -1:
            return None
        return remainder[1:end]
    for token in remainder.split():
        if token.startswith("\\"):
            continue
        return token.strip('"')
    return None


def _parse_hyperlink_url(instruction: str) -> str | None:
    """Return the URL argument from a ``HYPERLINK`` instruction.

    Accepts ``HYPERLINK "https://…"`` and ``HYPERLINK https://…``. Switches
    (``\\l`` for bookmark target, ``\\m``, ``\\n``, ``\\o``, ``\\t``) and their
    arguments are skipped; the first bare non-switch token is returned.
    """
    stripped = instruction.strip()
    if not stripped.upper().startswith("HYPERLINK"):
        return None
    remainder = stripped[len("HYPERLINK") :].strip()
    if not remainder:
        return None
    if remainder.startswith('"'):
        end = remainder.find('"', 1)
        if end == -1:
            return None
        return remainder[1:end]
    tokens = remainder.split()
    i = 0
    while i < len(tokens):
        token = tokens[i]
        if token.startswith("\\"):
            # -- switches `\l`, `\m`, `\n`, `\o`, `\t` take an argument --
            if token.lower() in (r"\l", r"\m", r"\n", r"\o", r"\t"):
                i += 2
                continue
            i += 1
            continue
        return token.strip('"')
    return None


def _tokenize_field_args(text: str) -> list[str]:
    """Tokenize `text` into whitespace-separated words, honouring double quotes
    and ``{...}`` nested-field groups as atomic tokens.

    ``{MERGEFIELD foo}`` is returned as the single token ``{MERGEFIELD foo}``
    (braces included) so the caller can spot nested fields. ``"quoted value"``
    returns ``quoted value`` (quotes stripped). Backslash-prefixed switches are
    returned verbatim including the backslash.
    """
    tokens: list[str] = []
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch.isspace():
            i += 1
            continue
        if ch == '"':
            end = text.find('"', i + 1)
            if end == -1:
                tokens.append(text[i + 1 :])
                break
            tokens.append(text[i + 1 : end])
            i = end + 1
            continue
        if ch == "{":
            depth = 1
            j = i + 1
            while j < n and depth > 0:
                if text[j] == "{":
                    depth += 1
                elif text[j] == "}":
                    depth -= 1
                j += 1
            tokens.append(text[i:j])
            i = j
            continue
        # -- bare word: read until whitespace, quote, or brace --
        j = i
        while j < n and not text[j].isspace() and text[j] not in ('"', "{"):
            j += 1
        tokens.append(text[i:j])
        i = j
    return tokens


def _resolve_operand(token: str, context: "dict[str, object]") -> str:
    """Resolve an operand token from :func:`_tokenize_field_args` into a string.

    * ``{MERGEFIELD name}`` or ``{ MERGEFIELD name }`` → ``context[name]`` as
      string (empty string when missing).
    * Anything else → returned verbatim (quoted literals are already
      unquoted by the tokenizer).
    """
    stripped = token.strip()
    if stripped.startswith("{") and stripped.endswith("}"):
        inner = stripped[1:-1].strip()
        if inner.upper().startswith("MERGEFIELD"):
            name = _parse_mergefield_name(inner)
            if name is None:
                return ""
            value = context.get(name)
            return "" if value is None else str(value)
        # -- unknown nested field: return empty string --
        return ""
    return token


def _compare(lhs: str, op: str, rhs: str) -> bool:
    """Compare two operands using `op`. Numeric comparison when both parse
    as ``float``; case-sensitive string compare otherwise.
    """
    try:
        lhs_n: float | None = float(lhs)
    except ValueError:
        lhs_n = None
    try:
        rhs_n: float | None = float(rhs)
    except ValueError:
        rhs_n = None

    if lhs_n is not None and rhs_n is not None:
        left: object = lhs_n
        right: object = rhs_n
    else:
        left = lhs
        right = rhs

    if op == "=":
        return left == right
    if op in ("<>", "!="):
        return left != right
    if op == "<":
        return left < right  # type: ignore[operator]
    if op == ">":
        return left > right  # type: ignore[operator]
    if op == "<=":
        return left <= right  # type: ignore[operator]
    if op == ">=":
        return left >= right  # type: ignore[operator]
    # -- unknown operator → falsy --
    return False


def _evaluate_if(
    instruction: str, context: "dict[str, object]"
) -> str | None:
    """Evaluate an ``IF`` field instruction against `context`.

    Expected shape (per ECMA-376 § 17.16.5.22):
    ``IF <expr1> <op> <expr2> <true-text> <false-text>``
    Returns the chosen text, or |None| when the instruction is malformed.
    """
    stripped = instruction.strip()
    if not stripped.upper().startswith("IF"):
        return None
    remainder = stripped[len("IF") :].strip()
    tokens = _tokenize_field_args(remainder)
    if len(tokens) < 4:
        return None
    # -- filter out formatting switches that may trail the field --
    core: list[str] = []
    skip_next = False
    for tok in tokens:
        if skip_next:
            skip_next = False
            continue
        if tok.startswith("\\"):
            if tok in ("\\*", "\\@", "\\#", "\\f"):
                skip_next = True
            continue
        core.append(tok)
    if len(core) < 4:
        return None
    lhs_raw, op, rhs_raw, true_text = core[0], core[1], core[2], core[3]
    false_text = core[4] if len(core) >= 5 else ""

    lhs = _resolve_operand(lhs_raw, context)
    rhs = _resolve_operand(rhs_raw, context)
    return true_text if _compare(lhs, op, rhs) else false_text


_ALLOWED_FORMULA_CHARS = set("0123456789.+-*/%() \t")

# -- AST node types permitted in ``=`` field expressions. Exponentiation
# -- (``ast.Pow``) is deliberately omitted: the per-character whitelist
# -- accepts two adjacent ``*`` chars, so ``9**9**9**9`` would otherwise
# -- slip through and pin CPU doing bignum arithmetic. --
_ALLOWED_FORMULA_NODES: tuple[type, ...] = (
    ast.Expression,
    ast.BinOp,
    ast.UnaryOp,
    ast.Constant,
)
# -- ``ast.Num`` was the legacy node type for numeric literals on Py<3.8;
# -- since 3.8 ``ast.parse`` emits ``ast.Constant`` for every literal, and
# -- the only thing ``ast.Num`` itself offers is a deprecated alias that
# -- emits ``DeprecationWarning`` on attribute access (Py 3.12+) and is
# -- removed entirely in Py 3.14. Our compat floor is 3.9, so
# -- ``ast.Constant`` already covers every literal we'll encounter and
# -- we don't need to whitelist the alias. --
_ALLOWED_FORMULA_OPS = (
    ast.Add,
    ast.Sub,
    ast.Mult,
    ast.Div,
    ast.Mod,
    ast.USub,
    ast.UAdd,
)


def _evaluate_formula_ast(expr: str) -> "int | float | None":
    """Safely evaluate `expr` as a small arithmetic expression.

    Returns the numeric result, or |None| if the expression contains any
    construct not in :data:`_ALLOWED_FORMULA_NODES` /
    :data:`_ALLOWED_FORMULA_OPS`. Rejects ``ast.Pow`` explicitly so the
    classic ``9**9**9**9`` bignum-DoS can never occur, regardless of what
    the character-class pre-filter accepts.
    """
    try:
        tree = ast.parse(expr, mode="eval")
    except (SyntaxError, ValueError):
        return None

    for node in ast.walk(tree):
        if isinstance(node, _ALLOWED_FORMULA_NODES):
            pass
        elif isinstance(node, _ALLOWED_FORMULA_OPS):
            pass
        else:
            return None

    def _eval(node):  # -- type: ignore[no-untyped-def] --
        if isinstance(node, ast.Expression):
            return _eval(node.body)
        if isinstance(node, ast.Constant):
            if isinstance(node.value, (int, float)):
                return node.value
            return None
        # -- ``ast.Num`` (legacy numeric-literal node) is unreachable on
        # -- Py 3.9+: ``ast.parse`` only emits ``ast.Constant``. The alias
        # -- itself emits ``DeprecationWarning`` on attribute access on
        # -- Py 3.12+ and is removed in 3.14, so we don't reference it. --
        if isinstance(node, ast.UnaryOp):
            operand = _eval(node.operand)
            if operand is None:
                return None
            if isinstance(node.op, ast.USub):
                return -operand
            if isinstance(node.op, ast.UAdd):
                return +operand
            return None
        if isinstance(node, ast.BinOp):
            left = _eval(node.left)
            right = _eval(node.right)
            if left is None or right is None:
                return None
            try:
                if isinstance(node.op, ast.Add):
                    return left + right
                if isinstance(node.op, ast.Sub):
                    return left - right
                if isinstance(node.op, ast.Mult):
                    return left * right
                if isinstance(node.op, ast.Div):
                    return left / right
                if isinstance(node.op, ast.Mod):
                    return left % right
            except (ZeroDivisionError, ArithmeticError, ValueError):
                return None
            return None
        return None

    return _eval(tree)


def _evaluate_formula(
    instruction: str, context: "dict[str, object]"
) -> str | None:
    """Evaluate a ``=`` field instruction as a restricted arithmetic expression.

    Nested ``{MERGEFIELD name}`` references are substituted from `context`
    first, then the resulting expression is checked against a small
    character-class whitelist (digits, the four basic operators, ``%``,
    parens, whitespace) *and* parsed to an :mod:`ast`, allowing only the
    node types in :data:`_ALLOWED_FORMULA_NODES`. Returns the string form
    of the result, or |None| on error.

    .. versionchanged:: 2026.05.11
       Replaced :func:`eval` with a whitelisted AST walker; ``ast.Pow``
       is explicitly rejected so the classic ``9**9**9**9`` bignum DoS
       cannot occur, regardless of what the character-class pre-filter
       accepts.
    """
    stripped = instruction.strip()
    if not stripped.startswith("="):
        return None
    expr = stripped[1:].strip()

    # -- substitute {MERGEFIELD name} → context[name] --
    def _sub(token: str) -> str:
        return _resolve_operand(token, context)

    parts: list[str] = []
    i = 0
    n = len(expr)
    while i < n:
        ch = expr[i]
        if ch == "{":
            depth = 1
            j = i + 1
            while j < n and depth > 0:
                if expr[j] == "{":
                    depth += 1
                elif expr[j] == "}":
                    depth -= 1
                j += 1
            parts.append(_sub(expr[i:j]))
            i = j
            continue
        parts.append(ch)
        i += 1
    substituted = "".join(parts)

    # -- whitelist check (retained as an inexpensive first line of defence) --
    for ch in substituted:
        if ch not in _ALLOWED_FORMULA_CHARS:
            return None
    if not substituted.strip():
        return None
    value = _evaluate_formula_ast(substituted)
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value)


# -- cross-references ----------------------------------------------------


_CROSS_REFERENCE_TYPES = frozenset(
    {"REF", "PAGEREF", "NOTEREF", "SEQREF", "STYLEREF"}
)


def build_cross_reference_instruction(
    ref_type: str,
    target_name: str,
    insert_as_hyperlink: bool = False,
    insert_paragraph_number: bool = False,
    insert_relative_position: bool = False,
    extra_switches: "list[str] | None" = None,
) -> str:
    """Return the field-code string for a cross-reference field.

    Builds the instruction shape used by Word for cross-reference complex
    fields:

    ``<REF_TYPE> <target_name> [\\h] [\\r] [\\p] [extra switches...]``

    The `target_name` is emitted verbatim if it is a well-formed bookmark
    name (letters, digits, underscore only); otherwise it is wrapped in
    double quotes so names with spaces round-trip cleanly. The caller may
    pass `extra_switches` as a list of raw switch tokens (e.g. ``["\\n"]``
    or ``["\\* MERGEFORMAT"]``) which are appended verbatim.

    .. versionadded:: 2026.05.10
    """
    ref_type = ref_type.strip().upper()
    if not ref_type:
        raise ValueError("ref_type must be a non-empty string")
    if not target_name:
        raise ValueError("target_name must be a non-empty string")

    # -- bookmark names in WordprocessingML may contain letters, digits, and
    #    underscore; anything else (spaces, punctuation) needs quoting --
    safe = all(ch.isalnum() or ch == "_" for ch in target_name)
    name_token = target_name if safe else f'"{target_name}"'

    parts: list[str] = [ref_type, name_token]
    if insert_as_hyperlink:
        parts.append("\\h")
    if insert_paragraph_number:
        parts.append("\\r")
    if insert_relative_position:
        parts.append("\\p")
    if extra_switches:
        parts.extend(extra_switches)
    return " ".join(parts)


class CrossReference(Field):
    """A :class:`Field` specialised for REF-family cross-reference fields.

    Returned by :meth:`Paragraph.add_cross_reference`. Also produced by
    :attr:`Field.as_cross_reference` for any existing |Field| whose type is
    one of ``REF``, ``PAGEREF``, ``NOTEREF``, ``SEQREF``, ``STYLEREF``.

    Exposes typed accessors for the cross-reference-specific pieces of the
    instruction — the reference type (e.g. ``PAGEREF``), the target name
    (typically a bookmark but may be a sequence identifier for ``SEQREF``
    or a style name for ``STYLEREF``), and the three ``\\h`` / ``\\r`` /
    ``\\p`` switches.

    Use :attr:`target_bookmark` with an owning :class:`Document` to resolve
    the named target to a |Bookmark| proxy (returns |None| when no matching
    bookmark exists — useful to detect broken cross-references).

    .. versionadded:: 2026.05.10
    """

    @property
    def ref_type(self) -> str:
        """The cross-reference type — ``"REF"``, ``"PAGEREF"``, etc.

        Alias for :attr:`type` with a domain-specific name. Always upper-case
        because field-type tokens are case-insensitive on input but canonical
        upper on output.

        .. versionadded:: 2026.05.10
        """
        return self.type.upper()

    @property
    def target_name(self) -> str:
        """The referenced target name — typically a bookmark name.

        For ``REF``, ``PAGEREF``, ``NOTEREF`` this is the bookmark name
        (the ``w:bookmarkStart/@w:name`` value a writer would match against).
        For ``SEQREF`` it is the sequence identifier. For ``STYLEREF`` it is
        the style name. Returns the empty string when the instruction has
        no target argument.

        .. versionadded:: 2026.05.10
        """
        name = _parse_ref_bookmark_name(self.instruction)
        return "" if name is None else name

    @property
    def insert_as_hyperlink(self) -> bool:
        """``True`` when the instruction carries the ``\\h`` switch.

        When set, Word renders the rendered result as a clickable link back
        to the referenced bookmark. Most authoring tools emit ``\\h`` for
        every cross-reference by default.

        .. versionadded:: 2026.05.10
        """
        return self._has_switch("h")

    @property
    def insert_paragraph_number(self) -> bool:
        """``True`` when the instruction carries the ``\\r`` switch.

        Requests the paragraph number of the referenced bookmark's paragraph
        (relative to the outline). Applicable to ``REF`` only in practice.

        .. versionadded:: 2026.05.10
        """
        return self._has_switch("r")

    @property
    def insert_relative_position(self) -> bool:
        """``True`` when the instruction carries the ``\\p`` switch.

        Requests the relative position ("above" / "below") of the target.
        Applicable to ``REF`` only in practice.

        .. versionadded:: 2026.05.10
        """
        return self._has_switch("p")

    def target_bookmark(self, document: "Document") -> "object | None":
        """Return the |Bookmark| proxy for the cross-reference's target.

        Looks up :attr:`target_name` in ``document.bookmarks`` and returns
        the matching |Bookmark|, or |None| when no bookmark with that name
        exists (broken cross-reference). Intended for ``REF``, ``PAGEREF``,
        and ``NOTEREF`` field types; for ``SEQREF`` and ``STYLEREF`` the
        "target" is a sequence / style and no |Bookmark| will be found.

        .. versionadded:: 2026.05.10
        """
        name = self.target_name
        if not name:
            return None
        return document.bookmarks.get(name)

    # -- internals ---------------------------------------------------------

    def _has_switch(self, letter: str) -> bool:
        """Return ``True`` if the instruction contains a ``\\<letter>`` switch.

        Letter comparison is case-insensitive (``\\h`` and ``\\H`` match).
        """
        parsed = parse_field_instruction(self.instruction)
        key = letter.upper()
        return key in parsed.switches


def _wrap_as_cross_reference(field: "Field") -> "CrossReference | None":
    """Return a :class:`CrossReference` view of `field`, or |None|.

    Returns |None| when `field` is not a cross-reference type. The returned
    object shares the same underlying element — mutations propagate.
    """
    if field.type.upper() not in _CROSS_REFERENCE_TYPES:
        return None
    return CrossReference(field._kind, field._element)  # pyright: ignore[reportPrivateUsage]


# -- Table-of-contents family -------------------------------------------


_TOC_FIELD_TYPES = frozenset({"TOC", "TOA", "TOF"})

# -- switches that take an argument in a TOC / TOA / TOF instruction,
#    per ECMA-376 § 17.16.5.68 (TOC), § 17.16.5.69 (TOA), § 17.16.5.70
#    (TOF). These override the generic `_ARG_TAKING_SWITCHES` for
#    field-specific parsing — in a REF field `\p` is a flag, in a TOC
#    field `\p` takes a string argument (the separator). --
_TOC_ARG_TAKING_SWITCHES = frozenset(
    {
        "a",  # \a: category abbreviation (TOA)
        "b",  # \b: bookmark name (TOC)
        "c",  # \c: caption label (TOC) / category number (TOA)
        "d",  # \d: separator between seq and page (TOC)
        "e",  # \e: entry separator (TOC with TC fields)
        "f",  # \f: TC-field ident filter (TOC)
        "g",  # \g: TOA sequence separator
        "l",  # \l: TC-field level range (TOC) / TOA levels
        "n",  # \n: omit-page-numbers range (TOC)
        "o",  # \o: outline-level range (TOC)
        "p",  # \p: separator (TOC) / passim text (TOA)
        "s",  # \s: sequence identifier (TOC)
        "t",  # \t: custom-style list (TOC)
    }
)


def parse_toc_instruction(instruction: str) -> ParsedFieldInstruction:
    """Parse a TOC / TOA / TOF instruction using TOC-specific switch rules.

    The generic :func:`parse_field_instruction` treats only ``\\*``, ``\\@``,
    ``\\#``, ``\\f`` as argument-taking. TOC fields extend the
    argument-taking set to ``\\a``, ``\\b``, ``\\c``, ``\\d``, ``\\e``,
    ``\\f``, ``\\g``, ``\\l``, ``\\n``, ``\\o``, ``\\p``, ``\\s``, ``\\t``.
    This parser honours the extended set so a switch-value like
    ``\\o "1-3"`` round-trips as ``switches["O"] == "1-3"`` instead of
    spilling into positional arguments.

    The ``\\n`` switch is special: it may appear with or without an
    argument (Word's "suppress page numbers everywhere" form is a bare
    ``\\n``). When the next token starts with a backslash, ``\\n`` is
    treated as a flag and recorded as ``switches["N"] == ""``.

    .. versionadded:: 2026.05.10
    """
    stripped = instruction.strip()
    if not stripped:
        return ParsedFieldInstruction(name="", args=[], switches={})

    tokens = _tokenize_field_args(stripped)
    if not tokens:
        return ParsedFieldInstruction(name="", args=[], switches={})

    name = tokens[0]
    args: list[str] = []
    switches: dict[str, str] = {}

    i = 1
    while i < len(tokens):
        token = tokens[i]
        if token.startswith("\\") and len(token) >= 2:
            raw_key = token[1:]
            key = raw_key.upper() if raw_key.isalpha() else raw_key
            tag = raw_key.lower() if raw_key.isalpha() else raw_key
            # -- TOC-scoped arg-taking switches first; fall back to the
            #    generic set (covers \* \@ \# \f) --
            takes_arg = tag in _TOC_ARG_TAKING_SWITCHES or tag in _ARG_TAKING_SWITCHES
            if takes_arg and i + 1 < len(tokens):
                next_token = tokens[i + 1]
                # -- special: \n is argument-taking only when followed by
                #    a non-switch token; a bare \n is the "all levels" form --
                if tag == "n" and next_token.startswith("\\"):
                    switches[key] = ""
                    i += 1
                    continue
                switches[key] = next_token
                i += 2
                continue
            switches[key] = ""
            i += 1
            continue
        args.append(token)
        i += 1

    return ParsedFieldInstruction(name=name, args=args, switches=switches)


def _parse_level_range(spec: str) -> "tuple[int, int] | None":
    """Return ``(min, max)`` parsed from a ``"1-3"`` / ``"2-2"`` range spec.

    Returns |None| when the spec is not a well-formed ``"<n>-<n>"`` string
    with ``1 <= min <= max <= 9``. Used for the ``\\o`` switch of a ``TOC``
    instruction and the ``\\n`` "omit page numbers for this range" switch.
    """
    if not spec or "-" not in spec:
        return None
    lo_s, _, hi_s = spec.partition("-")
    lo_s = lo_s.strip()
    hi_s = hi_s.strip()
    if not lo_s.isdigit() or not hi_s.isdigit():
        return None
    lo, hi = int(lo_s), int(hi_s)
    if not (1 <= lo <= hi <= 9):
        return None
    return (lo, hi)


def _parse_custom_styles(spec: str) -> "list[tuple[str, int]]":
    """Return ``[(style_name, level), ...]`` parsed from a ``\\t`` switch arg.

    Word encodes custom style mappings in the ``\\t`` switch as a
    comma-separated list of alternating style-name / level tokens, e.g.
    ``"Quote,1,Intense Quote,2"``. Returns an empty list when `spec` is
    empty, has no commas, or has an odd number of tokens.

    Malformed pair-level tokens (non-numeric) are skipped rather than
    raising. The caller can detect a fully-malformed value by checking for
    an empty return.
    """
    if not spec:
        return []
    tokens = [t.strip() for t in spec.split(",")]
    if len(tokens) < 2 or len(tokens) % 2 == 1:
        return []
    pairs: list[tuple[str, int]] = []
    for i in range(0, len(tokens), 2):
        style_name = tokens[i]
        level_s = tokens[i + 1]
        if not style_name or not level_s.isdigit():
            continue
        pairs.append((style_name, int(level_s)))
    return pairs


def _format_custom_styles(pairs: "list[tuple[str, int]]") -> str:
    """Return the ``"style,level,style,level"`` string for `pairs`.

    The inverse of :func:`_parse_custom_styles`. Style names are emitted
    verbatim; commas inside a style name are not escaped (Word doesn't
    support them either — the TOC dialog forbids commas in style names).
    """
    parts: list[str] = []
    for style_name, level in pairs:
        parts.append(f"{style_name},{level}")
    return ",".join(parts)


def build_toc_field_instruction(
    field_type: str = "TOC",
    heading_range: "tuple[int, int] | None" = (1, 3),
    hyperlinks: bool = True,
    hide_in_web: bool = True,
    use_outline_levels: bool = True,
    omit_page_numbers_range: "tuple[int, int] | None" = None,
    separator: "str | None" = None,
    custom_styles: "list[tuple[str, int]] | None" = None,
    caption_label: "str | None" = None,
    bookmark_name: "str | None" = None,
    extra_switches: "list[str] | None" = None,
) -> str:
    """Return the raw field-code string for a TOC-family field.

    `field_type` is one of ``"TOC"``, ``"TOA"``, ``"TOF"`` (case-insensitive
    on input; canonicalised to upper on output). The remaining arguments map
    to the ECMA-376 § 17.16.5.68 switch set:

    * `heading_range` → ``\\o "min-max"`` (outline-level range)
    * `hyperlinks` → ``\\h`` (render entries as hyperlinks)
    * `hide_in_web` → ``\\z`` (hide tab leader / page numbers in web layout)
    * `use_outline_levels` → ``\\u`` (use applied outline level on any
      paragraph, not just ``Heading N`` styles)
    * `omit_page_numbers_range` → ``\\n "min-max"`` (suppress page numbers
      for this outline-level range; passing ``(0, 0)`` emits a bare ``\\n``
      meaning "omit page numbers for all levels")
    * `separator` → ``\\p "sep"`` (character between entry text and page
      number — default is a tab)
    * `custom_styles` → ``\\t "style1,level1,style2,level2"`` (map custom
      paragraph styles to TOC levels)
    * `caption_label` → ``\\c "label"`` (include caption entries with this
      label — used by List of Figures / List of Tables)
    * `bookmark_name` → ``\\b "name"`` (restrict the TOC to headings inside
      this bookmark range)
    * `extra_switches` → appended verbatim (e.g. ``["\\w"]`` to preserve
      tab entries, or ``["\\x"]`` to preserve newline entries)

    The returned string is not wrapped in leading/trailing spaces; callers
    that append it into a ``w:instrText`` element typically add their own
    single-space padding to match Word's on-disk form.

    .. versionadded:: 2026.05.10
    """
    field_type = (field_type or "").strip().upper()
    if field_type not in _TOC_FIELD_TYPES:
        raise ValueError(
            "field_type must be 'TOC', 'TOA', or 'TOF'; got %r" % (field_type,)
        )

    parts: list[str] = [field_type]

    if heading_range is not None:
        lo, hi = heading_range
        if not (1 <= lo <= hi <= 9):
            raise ValueError(
                "heading_range must satisfy 1 <= min <= max <= 9, got %r"
                % (heading_range,)
            )
        parts.append(f'\\o "{lo}-{hi}"')

    if hyperlinks:
        parts.append("\\h")
    if hide_in_web:
        parts.append("\\z")
    if use_outline_levels:
        parts.append("\\u")

    if omit_page_numbers_range is not None:
        lo, hi = omit_page_numbers_range
        if lo == 0 and hi == 0:
            parts.append("\\n")
        else:
            if not (1 <= lo <= hi <= 9):
                raise ValueError(
                    "omit_page_numbers_range must satisfy 1 <= min <= max <= 9 "
                    "or be (0, 0), got %r" % (omit_page_numbers_range,)
                )
            parts.append(f'\\n "{lo}-{hi}"')

    if separator is not None:
        parts.append(f'\\p "{separator}"')

    if custom_styles:
        parts.append(f'\\t "{_format_custom_styles(custom_styles)}"')

    if caption_label is not None:
        parts.append(f'\\c "{caption_label}"')

    if bookmark_name is not None:
        parts.append(f'\\b "{bookmark_name}"')

    if extra_switches:
        parts.extend(extra_switches)

    return " ".join(parts)


class TocField(Field):
    """A :class:`Field` specialised for the ``TOC``-family fields.

    Returned by :meth:`Paragraph.add_toc`. Also produced by
    :attr:`Field.as_toc` for any existing |Field| whose type is one of
    ``TOC``, ``TOA``, ``TOF``. Two subclasses refine the base:

    * :class:`TableOfFiguresField` — a ``TOC`` with a ``\\c "label"`` switch
      (the shape Word emits for a *List of Figures* / *List of Tables*).
    * :class:`TableOfAuthoritiesField` — a ``TOA`` field (table of
      authorities).

    Exposes typed accessors for the most-common switches on TOC
    instructions:

    * :attr:`heading_range` → the ``\\o "min-max"`` outline-level range
    * :attr:`hyperlinks_enabled` → ``\\h``
    * :attr:`hide_in_web` → ``\\z``
    * :attr:`use_outline_levels` → ``\\u``
    * :attr:`omit_page_numbers_range` → ``\\n "min-max"`` (or |None| when
      the ``\\n`` switch is absent, or ``(0, 0)`` when ``\\n`` is bare)
    * :attr:`separator` → ``\\p "..."``
    * :attr:`custom_styles` → parsed ``\\t`` mapping
    * :attr:`caption_label` → ``\\c "..."``
    * :attr:`bookmark_name` → ``\\b "..."``

    .. versionadded:: 2026.05.10
    """

    @property
    def heading_range(self) -> "tuple[int, int] | None":
        """The ``(min, max)`` heading range from the ``\\o`` switch, or |None|.

        Returns |None| when the ``\\o`` switch is absent or malformed.
        """
        parsed = parse_toc_instruction(self.instruction)
        arg = parsed.switches.get("O")
        if arg is None:
            return None
        return _parse_level_range(arg)

    @property
    def hyperlinks_enabled(self) -> bool:
        """``True`` when the instruction carries the ``\\h`` switch."""
        return self._has_switch("H")

    @property
    def hide_in_web(self) -> bool:
        """``True`` when the instruction carries the ``\\z`` switch."""
        return self._has_switch("Z")

    @property
    def use_outline_levels(self) -> bool:
        """``True`` when the instruction carries the ``\\u`` switch."""
        return self._has_switch("U")

    @property
    def omit_page_numbers_range(self) -> "tuple[int, int] | None":
        """The ``(min, max)`` range from the ``\\n`` switch, or |None|.

        Returns ``(0, 0)`` when the ``\\n`` switch is present without an
        argument (Word's shorthand for "suppress page numbers for every
        level"). Returns |None| when the switch is absent.
        """
        parsed = parse_toc_instruction(self.instruction)
        if "N" not in parsed.switches:
            return None
        arg = parsed.switches["N"]
        if not arg:
            return (0, 0)
        parsed_range = _parse_level_range(arg)
        if parsed_range is None:
            return (0, 0)
        return parsed_range

    @property
    def separator(self) -> "str | None":
        """The ``\\p`` separator argument, or |None| when not set."""
        parsed = parse_toc_instruction(self.instruction)
        arg = parsed.switches.get("P")
        if arg is None or arg == "":
            return None
        return arg

    @property
    def custom_styles(self) -> "list[tuple[str, int]]":
        """The ``\\t`` custom-style mapping, or an empty list.

        Each tuple is ``(style_name, level)``. The list is empty when the
        ``\\t`` switch is absent or malformed.
        """
        parsed = parse_toc_instruction(self.instruction)
        arg = parsed.switches.get("T")
        if not arg:
            return []
        return _parse_custom_styles(arg)

    @property
    def caption_label(self) -> "str | None":
        """The ``\\c`` caption-label argument, or |None| when not set.

        Non-empty when the field is a List of Figures / List of Tables
        (``TOC \\c "Figure"`` / ``TOC \\c "Table"``).
        """
        parsed = parse_toc_instruction(self.instruction)
        arg = parsed.switches.get("C")
        if arg is None or arg == "":
            return None
        return arg

    @property
    def bookmark_name(self) -> "str | None":
        """The ``\\b`` bookmark-name argument, or |None| when not set."""
        parsed = parse_toc_instruction(self.instruction)
        arg = parsed.switches.get("B")
        if arg is None or arg == "":
            return None
        return arg

    # -- internals ---------------------------------------------------------

    def _has_switch(self, letter: str) -> bool:
        """Return ``True`` when the instruction carries a ``\\<letter>`` switch.

        Matches :class:`CrossReference._has_switch` — case-insensitive on
        the switch letter. Uses :func:`parse_toc_instruction` so switches
        that take arguments in a TOC context (``\\o``, ``\\n``, ``\\p``,
        ``\\t``, ``\\c``, ``\\b``, ``\\s``) are correctly distinguished
        from plain flags.
        """
        parsed = parse_toc_instruction(self.instruction)
        return letter.upper() in parsed.switches

    # -- content rebuild ---------------------------------------------------

    def rebuild(self, page_number_placeholder: str = "?") -> str:
        """Recompute the TOC's cached result from the document's headings.

        Walks the owning ``w:body`` ancestor, collects every paragraph whose
        style matches ``"Heading N"`` (case-insensitive) with ``N`` inside
        this field's :attr:`heading_range` (defaulting to ``(1, 9)`` when
        the ``\\o`` switch is absent), and writes a tab-separated preview
        between the ``separate`` and ``end`` markers — one line per heading,
        ``"{heading text}\\t{page_number_placeholder}"``.

        `page_number_placeholder` replaces the real page number that Word
        would compute after layout. This is **unavoidable**: python-docx has
        no layout engine, so accurate page numbers cannot be produced here.
        The default ``"?"`` matches the placeholder Word itself shows for a
        dirty TOC before the first refresh, which is exactly the state the
        field is in after :meth:`Paragraph.add_toc` — `mark_dirty` is
        preserved so Word recomputes real numbers on open.

        Returns the newly-written cached result string (empty when the
        document has no qualifying headings).

        Subclasses override this to redefine "what counts as a heading" —
        see :meth:`TableOfFiguresField.rebuild` for the caption-based variant.

        .. versionadded:: 2026.05.10
        """
        entries = self._collect_entries()
        result = self._format_entries(entries, page_number_placeholder)
        self.update_result_text(result)
        return result

    # -- internals: rebuild helpers ----------------------------------------

    def _collect_entries(self) -> "list[tuple[int, str]]":
        """Return ``(level, text)`` pairs for headings targeted by this TOC.

        Default implementation scans the document body for paragraphs
        styled ``"Heading N"`` with ``N`` inside :attr:`heading_range`
        (defaults to ``(1, 9)`` when no ``\\o`` switch is present).
        """
        heading_range = self.heading_range or (1, 9)
        min_level, max_level = heading_range
        entries: "list[tuple[int, str]]" = []
        for p in self._iter_body_paragraphs():
            level = _paragraph_heading_level(p)
            if level is None:
                continue
            if level < min_level or level > max_level:
                continue
            text = _paragraph_text(p)
            if not text:
                continue
            entries.append((level, text))
        return entries

    @staticmethod
    def _format_entries(
        entries: "list[tuple[int, str]]", placeholder: str
    ) -> str:
        """Join ``(level, text)`` entries into a tab-separated preview."""
        return "\n".join(f"{text}\t{placeholder}" for _, text in entries)

    def _iter_body_paragraphs(self):
        """Yield every ``w:p`` descendant of the owning ``w:body``, in order."""
        body = self._find_body_ancestor()
        if body is None:
            return
        for p in body.xpath(".//w:p"):
            yield p

    def _find_body_ancestor(self):
        """Return the nearest ``w:body`` ancestor of this field, or |None|."""
        ancestor = self._element.getparent()
        body_tag = qn("w:body")
        while ancestor is not None and ancestor.tag != body_tag:
            ancestor = ancestor.getparent()
        return ancestor


class TableOfFiguresField(TocField):
    """A :class:`TocField` specialised for ``TOC \\c "label"`` fields.

    Produced by :meth:`Paragraph.add_table_of_figures` and by
    :attr:`Field.as_toc` for any ``TOC`` field that carries a ``\\c``
    switch. The :attr:`caption_label` property on the base class already
    exposes the label; this subclass exists so callers can
    ``isinstance(field, TableOfFiguresField)`` to branch on *List of
    Figures* vs. *Table of Contents*.

    .. versionadded:: 2026.05.10
    """

    def _collect_entries(self) -> "list[tuple[int, str]]":
        """Return ``(0, text)`` pairs for caption paragraphs of this label.

        Overrides :meth:`TocField._collect_entries` to match the *List of
        Figures / Tables* semantics: every paragraph styled ``"Caption"``
        (case-insensitive) whose text begins with the ``\\c`` caption
        label (e.g. ``"Figure"`` or ``"Table"``) contributes an entry. The
        level is reported as ``0`` since a list of figures is flat.
        """
        label = self.caption_label
        entries: "list[tuple[int, str]]" = []
        for p in self._iter_body_paragraphs():
            if not _paragraph_is_caption(p):
                continue
            text = _paragraph_text(p)
            if not text:
                continue
            if label is not None and not text.lower().startswith(
                label.lower()
            ):
                continue
            entries.append((0, text))
        return entries


class TableOfAuthoritiesField(TocField):
    """A :class:`TocField` specialised for the ``TOA`` field type.

    Produced by :meth:`Paragraph.add_table_of_authorities` and by
    :attr:`Field.as_toc` for any ``TOA`` field. The TOA field shares the
    ``\\h``, ``\\c`` (category number for TOA), and ``\\b`` switches with
    TOC but has its own ``\\e``, ``\\g``, ``\\l``, ``\\p``, ``\\s`` forms
    that python-docx does not currently parse — callers that need those
    values can read :attr:`instruction` directly.

    .. versionadded:: 2026.05.10
    """

    @property
    def category(self) -> "int | None":
        """The TOA ``\\c`` category number, or |None| when absent.

        On a TOA field, ``\\c "N"`` selects which category (1 = cases,
        2 = statutes, …) to include. This differs from ``TOC \\c`` where
        the argument is a caption-label string.
        """
        parsed = parse_toc_instruction(self.instruction)
        arg = parsed.switches.get("C")
        if arg is None or arg == "":
            return None
        try:
            return int(arg)
        except ValueError:
            return None


def _wrap_as_toc(field: "Field") -> "TocField | None":
    """Return a :class:`TocField` (or subclass) view of `field`, or |None|.

    Dispatches to :class:`TableOfFiguresField` when the field is a ``TOC``
    with a ``\\c`` switch, :class:`TableOfAuthoritiesField` for ``TOA``,
    :class:`TocField` for plain ``TOC`` / ``TOF``. Returns |None| for
    everything else. The returned object shares the same underlying
    element — mutations propagate.
    """
    ft = field.type.upper()
    if ft not in _TOC_FIELD_TYPES:
        return None
    kind = field._kind  # pyright: ignore[reportPrivateUsage]
    elm = field._element  # pyright: ignore[reportPrivateUsage]
    if ft == "TOA":
        return TableOfAuthoritiesField(kind, elm)
    if ft == "TOC":
        parsed = parse_toc_instruction(field.instruction)
        if "C" in parsed.switches and parsed.switches["C"]:
            return TableOfFiguresField(kind, elm)
        return TocField(kind, elm)
    # -- TOF: bare "TOF" field type (rare; Word typically emits TOC \c
    #    instead, but ECMA-376 lists TOF as a separate field type too) --
    return TableOfFiguresField(kind, elm)


# -- TOC-rebuild helpers ---------------------------------------------------

# -- match both "Heading1" (the usual pStyle id Word emits for built-in
#    styles — whitespace stripped from the display name) and
#    "Heading 1" (the display name, which occasionally appears as the
#    pStyle val for documents authored by other tooling) --
_HEADING_STYLE_RE = re.compile(r"^heading\s*([1-9])$", re.IGNORECASE)


def _paragraph_style_name(p: "BaseOxmlElement") -> str:
    """Return the ``w:pPr/w:pStyle/@w:val`` of `p`, or the empty string."""
    vals = p.xpath("./w:pPr/w:pStyle/@w:val")
    if not vals:
        return ""
    return str(vals[0])


def _paragraph_heading_level(p: "BaseOxmlElement") -> "int | None":
    """Return the integer heading level of paragraph `p`, or |None|.

    Matches a ``pStyle`` whose ``w:val`` spells ``"Heading N"`` (the
    built-in English style IDs Word emits, regardless of the display
    name's locale). ``N`` must fall in 1..9.
    """
    name = _paragraph_style_name(p)
    if not name:
        return None
    match = _HEADING_STYLE_RE.match(name.strip())
    if match is None:
        return None
    return int(match.group(1))


def _paragraph_is_caption(p: "BaseOxmlElement") -> bool:
    """Return ``True`` when `p` is styled ``"Caption"`` (case-insensitive)."""
    name = _paragraph_style_name(p)
    return name.strip().lower() == "caption"


def _paragraph_text(p: "BaseOxmlElement") -> str:
    """Return the concatenated visible text of paragraph `p`.

    Joins every ``w:t`` descendant — matches the string
    :attr:`docx.text.paragraph.Paragraph.text` exposes, but avoids
    constructing a :class:`Paragraph` proxy (which requires a parent).
    """
    parts = p.xpath(".//w:t")
    return "".join(t.text or "" for t in parts)
