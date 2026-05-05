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

    .. versionadded:: 2026.05.0
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


def _evaluate_formula(
    instruction: str, context: "dict[str, object]"
) -> str | None:
    """Evaluate a ``=`` field instruction as a restricted arithmetic expression.

    Nested ``{MERGEFIELD name}`` references are substituted from `context`
    first, then the resulting expression is checked against a small
    character-class whitelist (digits, the four basic operators, ``%``,
    parens, whitespace) and evaluated with :func:`eval` in an empty
    namespace. Returns the string form of the result, or |None| on error.
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

    # -- whitelist check --
    for ch in substituted:
        if ch not in _ALLOWED_FORMULA_CHARS:
            return None
    if not substituted.strip():
        return None
    try:
        value = eval(substituted, {"__builtins__": {}}, {})  # noqa: S307
    except Exception:
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value)
