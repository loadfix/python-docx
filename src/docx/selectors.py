"""CSS-selector query language for python-docx documents.

A small, dependency-free CSS-selector parser and matcher that lets callers
locate proxy objects (paragraphs, runs, tables, rows, cells, hyperlinks,
bookmarks, comments) inside a |Document| with the familiar selector syntax
borrowed from the web::

    headings = doc.select('p[style="Heading 1"]')
    bold_in_headings = doc.select('p[style^="Heading "] r[bold]')
    light_tables = doc.select('tbl[style="Light List"]')
    second_col_cells = doc.select('tbl tr td:nth-child(2)')
    intros = doc.select('p[style^="Heading "] + p')

The selector engine deliberately speaks an OOXML-flavoured dialect rather
than reusing CSS verbatim — the element types, attribute names, and the
pseudo-class semantics are tuned to the python-docx proxy graph, not the
HTML/CSS DOM. Only the syntax is shared.

Cheatsheet
----------

Element types (the "tag" of the selector)::

    p            paragraph (``<w:p>``)
    r            run (``<w:r>``)
    tbl          table (``<w:tbl>``)
    tr           table row (``<w:tr>``)
    td           table cell (``<w:tc>``)
    hyperlink    hyperlink (``<w:hyperlink>``)
    bookmark     bookmark (``<w:bookmarkStart>`` + matching end)
    comment      comment (``CommentPart`` entries)

Attribute selectors (the bracketed clause)::

    [style=Heading 1]    exact-match style (id or display name)
    [style^=Heading ]    starts-with
    [style$=tail]        ends-with
    [style*=ead]         contains-substring
    [style]              has the attribute (style applied)
    [bold]               boolean attribute is True (also: italic,
                         underline, hidden — runs only)

Combinators::

    p r              descendant — runs inside paragraphs
    tbl > tr         child — direct ``tr`` children of a table
    p + p            adjacent sibling — paragraph immediately after
                     another paragraph

Pseudo-classes::

    :first-child     first matching element among its peers in the
                     enclosing scope
    :last-child      last matching element among its peers
    :nth-child(N)    1-indexed position; ``N`` may be a positive int
                     (``:nth-child(2)``) or the literal ``odd`` /
                     ``even`` keywords
    :not(simple)     negate a simple selector (one element-type plus
                     attributes / pseudo-classes — no combinators)

Compound selectors such as ``p.heading[level=1]`` chain a tag with any
number of attribute clauses, class/id shorthands (``.heading`` is sugar
for ``[class=heading]`` and ``#summary`` for ``[id=summary]``), and
pseudo-classes. Whitespace separates compounds; the ``>`` and ``+``
combinators must be flanked by whitespace.

Proxy types returned
--------------------

The matcher returns the same proxy objects callers already hold for each
element kind — :class:`docx.text.paragraph.Paragraph`,
:class:`docx.text.run.Run`, :class:`docx.table.Table`,
:class:`docx.table._Row`, :class:`docx.table._Cell`,
:class:`docx.text.hyperlink.Hyperlink`,
:class:`docx.bookmarks.Bookmark`, and
:class:`docx.comments.Comment`. Callers can use the result objects with
the regular python-docx surface (``run.bold = True``, ``p.text``, etc.).

Limitations
-----------

The engine ships an intentionally small CSS subset. Unsupported features
include the general-sibling combinator (``~``), attribute case-folding
flags (``[name=val i]``), function-form ``:nth-child(an+b)``, and
selector lists separated by ``,``. Selectors that mention element types
outside the eight supported above raise :class:`SelectorSyntaxError` so
typos surface early.

.. versionadded:: 2026.05.13
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import (
    TYPE_CHECKING,
    Any,
    Callable,
    Iterable,
    Iterator,
    List,
    Optional,
    Sequence,
    Tuple,
    Union,
)

if TYPE_CHECKING:
    from docx.document import Document


__all__ = [
    "SelectorSyntaxError",
    "compile_selector",
    "select",
    "select_one",
]


# -- Public exception ---------------------------------------------------------


class SelectorSyntaxError(ValueError):
    """Raised when a selector cannot be parsed.

    Carries the offending selector (``selector`` attribute) plus a short
    human-readable message describing what the parser was expecting.
    """

    def __init__(self, message: str, selector: str = ""):
        super().__init__(message)
        self.selector = selector


# -- AST nodes ----------------------------------------------------------------


@dataclass
class _AttrPredicate:
    """One ``[name op value]`` (or boolean ``[name]``) clause."""

    name: str
    op: Optional[str]  # None | "=" | "^=" | "$=" | "*=" | "exists"
    value: Optional[str] = None


@dataclass
class _PseudoPredicate:
    """One ``:pseudo-class[(arg)]`` clause."""

    name: str  # "first-child" | "last-child" | "nth-child" | "not"
    arg: Any = None  # int / "odd" / "even" / SimpleSelector


@dataclass
class _SimpleSelector:
    """``tag`` plus zero or more attribute / pseudo predicates."""

    tag: str
    attrs: List[_AttrPredicate] = field(default_factory=list)
    pseudos: List[_PseudoPredicate] = field(default_factory=list)


# Combinator: " " (descendant), ">" (child), "+" (adjacent sibling).
_Combinator = str  # one of " ", ">", "+"


@dataclass
class _CompoundSelector:
    """A chain of simple selectors joined by combinators.

    ``parts`` is a list of ``(combinator, simple)`` pairs. The first
    pair's combinator is meaningless and is set to ``" "`` for
    consistency.
    """

    parts: List[Tuple[_Combinator, _SimpleSelector]]


# -- Tokenizer / parser -------------------------------------------------------


_TAGS = frozenset({"p", "r", "tbl", "tr", "td", "hyperlink", "bookmark", "comment"})


_BOOL_RUN_ATTRS = frozenset({"bold", "italic", "underline", "hidden"})


_PSEUDOS = frozenset({"first-child", "last-child", "nth-child", "not"})


# Tokens recognised by the lexer. ``IDENT`` covers tag names, attribute
# names, and unquoted attribute values; quoted values are emitted as a
# single ``STRING`` token.
_TOKEN_RE = re.compile(
    r"""
    \s+                        # whitespace
    | (?P<gt>>)                 # > combinator
    | (?P<plus>\+)              # + combinator
    | (?P<lbracket>\[)          # [
    | (?P<rbracket>\])          # ]
    | (?P<lparen>\()            # (
    | (?P<rparen>\))            # )
    | (?P<colon>:)              # :
    | (?P<dot>\.)               # . class shorthand
    | (?P<hash>\#)              # # id shorthand
    | (?P<starts>\^=)           # ^=
    | (?P<ends>\$=)             # $=
    | (?P<contains>\*=)         # *=
    | (?P<eq>=)                 # =
    | (?P<string>"[^"]*"|'[^']*')  # quoted string
    | (?P<ident>[A-Za-z_][A-Za-z0-9_-]*)  # ident
    | (?P<number>-?\d+)         # signed integer
    """,
    re.VERBOSE,
)


def _tokenize(selector: str) -> List[Tuple[str, str]]:
    """Return a list of ``(kind, text)`` tokens for ``selector``.

    Whitespace between meaningful tokens is collapsed into a single
    ``WS`` token so the parser can treat it as the descendant
    combinator.
    """
    tokens: List[Tuple[str, str]] = []
    pos = 0
    n = len(selector)
    while pos < n:
        ch = selector[pos]
        if ch.isspace():
            # collapse runs of whitespace
            j = pos + 1
            while j < n and selector[j].isspace():
                j += 1
            tokens.append(("WS", " "))
            pos = j
            continue
        m = _TOKEN_RE.match(selector, pos)
        if m is None:
            raise SelectorSyntaxError(
                f"unexpected character {ch!r} at position {pos}", selector
            )
        kind = m.lastgroup
        if kind is None:
            raise SelectorSyntaxError(
                f"unexpected character {ch!r} at position {pos}", selector
            )
        text = m.group(kind)
        tokens.append((kind, text))
        pos = m.end()
    return tokens


class _Parser:
    """Recursive-descent parser over the lexer's token stream."""

    def __init__(self, selector: str):
        self.selector = selector
        # drop any leading whitespace tokens — the parser deals with WS
        # only when it actually delimits two simple selectors.
        raw = _tokenize(selector)
        # strip leading and trailing WS
        while raw and raw[0][0] == "WS":
            raw.pop(0)
        while raw and raw[-1][0] == "WS":
            raw.pop()
        self.tokens = raw
        self.pos = 0

    def _peek(self, offset: int = 0) -> Optional[Tuple[str, str]]:
        i = self.pos + offset
        if i >= len(self.tokens):
            return None
        return self.tokens[i]

    def _advance(self) -> Tuple[str, str]:
        tok = self.tokens[self.pos]
        self.pos += 1
        return tok

    def _expect(self, kind: str) -> Tuple[str, str]:
        tok = self._peek()
        if tok is None or tok[0] != kind:
            got = "EOF" if tok is None else f"{tok[0]} {tok[1]!r}"
            raise SelectorSyntaxError(
                f"expected {kind}, got {got}", self.selector
            )
        return self._advance()

    # -- top-level parse ----------------------------------------------------

    def parse(self) -> _CompoundSelector:
        if not self.tokens:
            raise SelectorSyntaxError("empty selector", self.selector)
        parts: List[Tuple[_Combinator, _SimpleSelector]] = []
        first = self._parse_simple()
        parts.append((" ", first))
        while self.pos < len(self.tokens):
            combinator = self._consume_combinator()
            simple = self._parse_simple()
            parts.append((combinator, simple))
        return _CompoundSelector(parts=parts)

    def _consume_combinator(self) -> _Combinator:
        tok = self._peek()
        if tok is None:
            raise SelectorSyntaxError(
                "unexpected end of selector", self.selector
            )
        if tok[0] == "WS":
            self._advance()
            # next token may be a > or + (allowing either spacing style)
            nxt = self._peek()
            if nxt and nxt[0] == "gt":
                self._advance()
                # eat optional trailing WS
                if self._peek() and self._peek()[0] == "WS":
                    self._advance()
                return ">"
            if nxt and nxt[0] == "plus":
                self._advance()
                if self._peek() and self._peek()[0] == "WS":
                    self._advance()
                return "+"
            return " "
        if tok[0] == "gt":
            self._advance()
            if self._peek() and self._peek()[0] == "WS":
                self._advance()
            return ">"
        if tok[0] == "plus":
            self._advance()
            if self._peek() and self._peek()[0] == "WS":
                self._advance()
            return "+"
        # No explicit combinator — treat as descendant.
        return " "

    # -- simple selector ----------------------------------------------------

    def _parse_simple(self) -> _SimpleSelector:
        tok = self._peek()
        if tok is None:
            raise SelectorSyntaxError(
                "expected a tag at start of compound selector",
                self.selector,
            )
        if tok[0] != "ident":
            raise SelectorSyntaxError(
                f"expected an element type, got {tok[1]!r}", self.selector
            )
        tag = self._advance()[1]
        if tag not in _TAGS:
            raise SelectorSyntaxError(
                f"unknown element type {tag!r}; supported: "
                f"{', '.join(sorted(_TAGS))}",
                self.selector,
            )
        simple = _SimpleSelector(tag=tag)
        while True:
            nxt = self._peek()
            if nxt is None:
                break
            if nxt[0] == "lbracket":
                simple.attrs.append(self._parse_attr())
                continue
            if nxt[0] == "dot":
                self._advance()
                ident = self._expect("ident")[1]
                simple.attrs.append(
                    _AttrPredicate(name="class", op="=", value=ident)
                )
                continue
            if nxt[0] == "hash":
                self._advance()
                ident = self._expect("ident")[1]
                simple.attrs.append(
                    _AttrPredicate(name="id", op="=", value=ident)
                )
                continue
            if nxt[0] == "colon":
                simple.pseudos.append(self._parse_pseudo())
                continue
            break
        return simple

    def _parse_attr(self) -> _AttrPredicate:
        self._expect("lbracket")
        name = self._expect("ident")[1]
        nxt = self._peek()
        if nxt is None:
            raise SelectorSyntaxError(
                "unterminated attribute selector", self.selector
            )
        if nxt[0] == "rbracket":
            self._advance()
            return _AttrPredicate(name=name, op="exists", value=None)
        op_map = {"eq": "=", "starts": "^=", "ends": "$=", "contains": "*="}
        if nxt[0] not in op_map:
            raise SelectorSyntaxError(
                f"expected operator inside [{name}...], got {nxt[1]!r}",
                self.selector,
            )
        op = op_map[self._advance()[0]]
        val_tok = self._peek()
        if val_tok is None:
            raise SelectorSyntaxError(
                "expected attribute value", self.selector
            )
        if val_tok[0] == "string":
            value = self._advance()[1][1:-1]
        elif val_tok[0] in ("ident", "number"):
            # Unquoted value — concatenate consecutive ident/number/space-free
            # tokens until we hit ``]`` so users can write ``[style=Heading 1]``
            # without quotes (the parser collapses interior whitespace into a
            # single space, matching how Word actually stores style names).
            parts: List[str] = [self._advance()[1]]
            while True:
                lookahead = self._peek()
                if lookahead is None:
                    raise SelectorSyntaxError(
                        "unterminated attribute selector", self.selector
                    )
                if lookahead[0] == "rbracket":
                    break
                if lookahead[0] == "WS":
                    self._advance()
                    parts.append(" ")
                    continue
                if lookahead[0] in ("ident", "number"):
                    parts.append(self._advance()[1])
                    continue
                raise SelectorSyntaxError(
                    f"unexpected token {lookahead[1]!r} in attribute value",
                    self.selector,
                )
            # collapse trailing space artefacts.
            value = "".join(parts).strip()
        else:
            raise SelectorSyntaxError(
                f"expected attribute value, got {val_tok[1]!r}",
                self.selector,
            )
        self._expect("rbracket")
        return _AttrPredicate(name=name, op=op, value=value)

    def _parse_pseudo(self) -> _PseudoPredicate:
        self._expect("colon")
        name_parts: List[str] = [self._expect("ident")[1]]
        # pseudo-class names are dash-delimited idents (e.g. ``first-child``).
        # The lexer preserves the dashes inside ident tokens, so a single
        # ident is enough.
        name = name_parts[0]
        if name not in _PSEUDOS:
            raise SelectorSyntaxError(
                f"unsupported pseudo-class :{name}", self.selector
            )
        arg: Any = None
        nxt = self._peek()
        if nxt and nxt[0] == "lparen":
            self._advance()
            if name == "nth-child":
                tok = self._peek()
                if tok is None:
                    raise SelectorSyntaxError(
                        ":nth-child needs an argument", self.selector
                    )
                if tok[0] == "number":
                    n = int(self._advance()[1])
                    if n < 1:
                        raise SelectorSyntaxError(
                            ":nth-child argument must be >= 1, got "
                            + str(n),
                            self.selector,
                        )
                    arg = n
                elif tok[0] == "ident" and tok[1] in ("odd", "even"):
                    arg = self._advance()[1]
                else:
                    raise SelectorSyntaxError(
                        f":nth-child argument {tok[1]!r} not "
                        "supported (use a positive integer or "
                        "'odd' / 'even')",
                        self.selector,
                    )
            elif name == "not":
                # The negation accepts a *simple* selector — recurse.
                # If the inner starts with an attribute / pseudo (no
                # tag), the negation is interpreted relative to the
                # outer compound's tag, matching CSS semantics
                # (``p:not(:first-child)``).
                tok = self._peek()
                if tok is not None and tok[0] in ("lbracket", "colon", "dot", "hash"):
                    inner_simple = _SimpleSelector(tag="*")
                    while True:
                        nxt = self._peek()
                        if nxt is None:
                            break
                        if nxt[0] == "lbracket":
                            inner_simple.attrs.append(self._parse_attr())
                            continue
                        if nxt[0] == "dot":
                            self._advance()
                            ident = self._expect("ident")[1]
                            inner_simple.attrs.append(
                                _AttrPredicate(name="class", op="=", value=ident)
                            )
                            continue
                        if nxt[0] == "hash":
                            self._advance()
                            ident = self._expect("ident")[1]
                            inner_simple.attrs.append(
                                _AttrPredicate(name="id", op="=", value=ident)
                            )
                            continue
                        if nxt[0] == "colon":
                            inner_simple.pseudos.append(self._parse_pseudo())
                            continue
                        break
                else:
                    inner_simple = self._parse_simple()
                arg = inner_simple
            else:
                raise SelectorSyntaxError(
                    f":{name} does not accept an argument",
                    self.selector,
                )
            self._expect("rparen")
        else:
            if name in ("nth-child", "not"):
                raise SelectorSyntaxError(
                    f":{name} requires a parenthesised argument",
                    self.selector,
                )
        return _PseudoPredicate(name=name, arg=arg)


def compile_selector(selector: str) -> _CompoundSelector:
    """Parse ``selector`` into a reusable AST.

    Raises :class:`SelectorSyntaxError` on malformed input. The returned
    object can be passed to :func:`select` / :func:`select_one` to skip
    the parse step on hot paths.
    """
    return _Parser(selector).parse()


# -- Universe / candidate gathering -------------------------------------------

# A lightweight wrapper used while matching: keeps the proxy alongside an
# integer position inside its enclosing scope (cell, paragraph, body) so
# pseudo-classes like ``:nth-child(2)`` can be evaluated without
# revisiting the underlying XML.


@dataclass
class _Candidate:
    proxy: Any
    parent: Any  # parent proxy (or Document for top-level)
    index: int  # 1-based index inside ``parent``'s peer collection
    siblings: int  # total peer count (so :last-child is cheap)
    tag: str  # echo of the element kind, e.g. "p" or "r"


def _gather_paragraphs(document: "Document") -> List[_Candidate]:
    """Yield one |_Candidate| per paragraph in the document body and tables.

    The ``parent`` is the enclosing |Document| (for body paragraphs) or
    the enclosing |_Cell| proxy (for table-cell paragraphs); ``index``
    is 1-based among that parent's paragraphs.
    """
    out: List[_Candidate] = []
    body_paragraphs = list(document.paragraphs)
    n_body = len(body_paragraphs)
    for i, p in enumerate(body_paragraphs, start=1):
        out.append(_Candidate(p, document, i, n_body, "p"))
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                cell_paragraphs = list(cell.paragraphs)
                n_cell = len(cell_paragraphs)
                for i, p in enumerate(cell_paragraphs, start=1):
                    out.append(_Candidate(p, cell, i, n_cell, "p"))
    return out


def _gather_runs(document: "Document") -> List[_Candidate]:
    """Yield |_Candidate| objects for every run reachable from the body."""
    out: List[_Candidate] = []
    for p_cand in _gather_paragraphs(document):
        runs = list(p_cand.proxy.runs)
        n = len(runs)
        for i, r in enumerate(runs, start=1):
            out.append(_Candidate(r, p_cand.proxy, i, n, "r"))
    return out


def _gather_tables(document: "Document") -> List[_Candidate]:
    out: List[_Candidate] = []
    body_tables = list(document.tables)
    n = len(body_tables)
    for i, t in enumerate(body_tables, start=1):
        out.append(_Candidate(t, document, i, n, "tbl"))
    return out


def _gather_rows(document: "Document") -> List[_Candidate]:
    out: List[_Candidate] = []
    for t_cand in _gather_tables(document):
        rows = list(t_cand.proxy.rows)
        n = len(rows)
        for i, r in enumerate(rows, start=1):
            out.append(_Candidate(r, t_cand.proxy, i, n, "tr"))
    return out


def _gather_cells(document: "Document") -> List[_Candidate]:
    out: List[_Candidate] = []
    for r_cand in _gather_rows(document):
        cells = list(r_cand.proxy.cells)
        n = len(cells)
        for i, c in enumerate(cells, start=1):
            out.append(_Candidate(c, r_cand.proxy, i, n, "td"))
    return out


def _gather_hyperlinks(document: "Document") -> List[_Candidate]:
    out: List[_Candidate] = []
    for p_cand in _gather_paragraphs(document):
        # Paragraph.hyperlinks is the public collection.
        try:
            links = list(p_cand.proxy.hyperlinks)
        except AttributeError:
            links = []
        n = len(links)
        for i, link in enumerate(links, start=1):
            out.append(_Candidate(link, p_cand.proxy, i, n, "hyperlink"))
    return out


def _gather_bookmarks(document: "Document") -> List[_Candidate]:
    bookmarks = list(document.bookmarks)
    out: List[_Candidate] = []
    n = len(bookmarks)
    for i, b in enumerate(bookmarks, start=1):
        out.append(_Candidate(b, document, i, n, "bookmark"))
    return out


def _gather_comments(document: "Document") -> List[_Candidate]:
    try:
        comments = list(document.comments)
    except Exception:
        comments = []
    out: List[_Candidate] = []
    n = len(comments)
    for i, c in enumerate(comments, start=1):
        out.append(_Candidate(c, document, i, n, "comment"))
    return out


_GATHERERS: dict[str, Callable[["Document"], List[_Candidate]]] = {
    "p": _gather_paragraphs,
    "r": _gather_runs,
    "tbl": _gather_tables,
    "tr": _gather_rows,
    "td": _gather_cells,
    "hyperlink": _gather_hyperlinks,
    "bookmark": _gather_bookmarks,
    "comment": _gather_comments,
}


# -- Attribute & pseudo evaluation --------------------------------------------


def _resolve_style_value(proxy: Any) -> Optional[str]:
    """Return the style display name of ``proxy`` if it carries one.

    Falls back to the style id when the style proxy does not expose a
    ``name`` (Word stores both; matching either lets users write either
    spelling in their selector).
    """
    style = getattr(proxy, "style", None)
    if style is None:
        return None
    name = getattr(style, "name", None)
    if name:
        return str(name)
    sid = getattr(style, "style_id", None)
    if sid:
        return str(sid)
    return None


def _attr_value(proxy: Any, name: str) -> Optional[Any]:
    """Best-effort attribute lookup over the proxy graph.

    The selector language uses CSS-style attribute names; the python-docx
    proxies expose them under several shapes (Pythonic ``style.name``,
    ``font.bold`` on runs, ``address`` / ``tooltip`` on hyperlinks, …).
    This helper centralises the mapping so the matching code stays small.
    """
    if name == "style":
        return _resolve_style_value(proxy)
    if name == "text":
        return getattr(proxy, "text", None)
    if name == "name":
        # bookmarks / comments expose a plain ``name`` attr
        return getattr(proxy, "name", None)
    if name == "id":
        for attr in ("bookmark_id", "comment_id"):
            v = getattr(proxy, attr, None)
            if v is not None:
                return str(v)
        v = getattr(proxy, "id", None)
        return None if v is None else str(v)
    if name == "address":
        return getattr(proxy, "address", None)
    if name == "tooltip":
        return getattr(proxy, "tooltip", None)
    if name == "author":
        return getattr(proxy, "author", None)
    if name == "level":
        # heading level extracted from the style name
        sval = _resolve_style_value(proxy)
        if sval and sval.startswith("Heading "):
            tail = sval[len("Heading "):].strip()
            if tail.isdigit():
                return tail
        return None
    if name == "class":
        # treat class as style display name (mirrors HTML's loose
        # mapping of class -> style)
        return _resolve_style_value(proxy)
    if name in _BOOL_RUN_ATTRS:
        # boolean run flags live behind ``run.bold`` / ``run.italic`` etc.
        return getattr(proxy, name, None)
    # last resort: a plain attribute lookup
    return getattr(proxy, name, None)


def _match_attr(proxy: Any, pred: _AttrPredicate) -> bool:
    if pred.op == "exists" or pred.op is None:
        # Boolean OOXML flag (e.g. ``[bold]``) — Word stores tri-states
        # so True is the only positive match.
        if pred.name in _BOOL_RUN_ATTRS:
            return getattr(proxy, pred.name, None) is True
        v = _attr_value(proxy, pred.name)
        return v is not None and v != ""
    actual = _attr_value(proxy, pred.name)
    if actual is None:
        return False
    actual_str = str(actual)
    needle = pred.value or ""
    if pred.op == "=":
        return actual_str == needle
    if pred.op == "^=":
        return actual_str.startswith(needle)
    if pred.op == "$=":
        return actual_str.endswith(needle)
    if pred.op == "*=":
        return needle in actual_str
    return False  # pragma: no cover -- parser rejects unknown ops first


def _match_pseudo(cand: _Candidate, pred: _PseudoPredicate) -> bool:
    name = pred.name
    if name == "first-child":
        return cand.index == 1
    if name == "last-child":
        return cand.index == cand.siblings
    if name == "nth-child":
        if isinstance(pred.arg, int):
            return cand.index == pred.arg
        if pred.arg == "odd":
            return cand.index % 2 == 1
        if pred.arg == "even":
            return cand.index % 2 == 0
        return False  # pragma: no cover
    if name == "not":
        inner: _SimpleSelector = pred.arg
        if inner.tag == "*":
            # Tag-less inner — match attrs/pseudos against `cand` as-is.
            for ap in inner.attrs:
                if not _match_attr(cand.proxy, ap):
                    return True
            for pp in inner.pseudos:
                if not _match_pseudo(cand, pp):
                    return True
            return False
        if inner.tag != cand.tag:
            return True  # different tag -> the negation passes
        return not _match_simple(cand, inner)
    return False  # pragma: no cover


def _match_simple(cand: _Candidate, simple: _SimpleSelector) -> bool:
    if simple.tag != cand.tag:
        return False
    for pred in simple.attrs:
        if not _match_attr(cand.proxy, pred):
            return False
    for pred in simple.pseudos:
        if not _match_pseudo(cand, pred):
            return False
    return True


# -- Combinator handling ------------------------------------------------------


def _is_descendant(child_proxy: Any, ancestor_proxy: Any) -> bool:
    """Return True when ``child_proxy`` is contained in ``ancestor_proxy``.

    Uses the underlying lxml ``_element`` for a quick ancestor walk —
    this is robust to the proxy types because every python-docx proxy
    exposes its CT_ element via ``_element`` (and the documents
    themselves expose ``_element`` returning the ``CT_Document`` root).
    """
    child_el = getattr(child_proxy, "_element", None)
    anc_el = getattr(ancestor_proxy, "_element", None)
    if child_el is None or anc_el is None:
        return False
    parent = child_el.getparent()
    while parent is not None:
        if parent is anc_el:
            return True
        parent = parent.getparent()
    return False


def _is_child(child_proxy: Any, parent_proxy: Any) -> bool:
    """True when ``child_proxy``'s logical parent is ``parent_proxy``.

    The candidate gatherers populate ``_Candidate.parent`` with the
    *logical* parent — body for top-level paragraphs, _Cell for cell
    paragraphs, _Row for cells, etc. Two distinct proxy instances may
    wrap the same underlying CT_ element, so we identity-check both
    the proxy itself and (when both expose ``_element``) the wrapped
    element to recognise the same logical container.
    """
    if child_proxy is parent_proxy:
        return True
    a = getattr(child_proxy, "_element", None)
    b = getattr(parent_proxy, "_element", None)
    return a is not None and b is not None and a is b


def _is_adjacent_sibling(prev_proxy: Any, next_proxy: Any) -> bool:
    """True when ``next_proxy`` immediately follows ``prev_proxy`` in
    document order at the same nesting level.
    """
    prev_el = getattr(prev_proxy, "_element", None)
    next_el = getattr(next_proxy, "_element", None)
    if prev_el is None or next_el is None:
        return False
    if prev_el.getparent() is not next_el.getparent():
        return False
    sibling = prev_el.getnext()
    # Skip any non-block siblings (bookmark markers, sectPr, etc.) to
    # the next "real" sibling that matches the second compound's element
    # kind. We only need to know whether ``next_el`` is the *first*
    # block-or-tag-equivalent sibling, so walking until we hit it is
    # sufficient.
    while sibling is not None:
        if sibling is next_el:
            return True
        # A run inside a paragraph counts every non-text run as a
        # sibling. A paragraph-paragraph adjacency must skip
        # ``w:bookmarkStart`` / ``w:bookmarkEnd`` / ``w:sectPr``.
        if sibling.tag.endswith("}p") or sibling.tag.endswith("}r") or sibling.tag.endswith("}tbl"):
            return False
        sibling = sibling.getnext()
    return False


# -- Main entry points --------------------------------------------------------


def _matches_compound(
    cand: _Candidate,
    compound: _CompoundSelector,
    document: "Document",
) -> bool:
    """Walk the compound selector right-to-left for ``cand``.

    Each step pops a (combinator, simple) pair off the right end and
    checks that some ancestor / parent / preceding-sibling candidate
    matches the next-simpler selector. The recursion bottoms out when
    only the leftmost simple selector is left — which is matched against
    the current candidate directly.
    """
    parts = compound.parts
    if not parts:
        return False
    # The rightmost simple selector must match this candidate.
    last_combinator, last_simple = parts[-1]
    if not _match_simple(cand, last_simple):
        return False
    if len(parts) == 1:
        return True
    # Walk left across the remaining parts; each one's `combinator`
    # describes how it relates to the part on its *right*.
    return _walk_left(cand, parts, len(parts) - 1, document)


def _walk_left(
    cand: _Candidate,
    parts: List[Tuple[_Combinator, _SimpleSelector]],
    idx: int,
    document: "Document",
) -> bool:
    """Recursive helper; ``parts[idx]`` already matched ``cand``."""
    if idx == 0:
        return True
    combinator = parts[idx][0]
    prev_simple = parts[idx - 1][1]
    # gather candidates of the prev_simple's tag.
    prev_candidates = _GATHERERS[prev_simple.tag](document)
    for pc in prev_candidates:
        if not _match_simple(pc, prev_simple):
            continue
        ok = False
        if combinator == " ":
            ok = _is_descendant(cand.proxy, pc.proxy)
        elif combinator == ">":
            ok = _is_child(cand.parent, pc.proxy)
        elif combinator == "+":
            ok = _is_adjacent_sibling(pc.proxy, cand.proxy)
        if ok and _walk_left(pc, parts, idx - 1, document):
            return True
    return False


def select(document: "Document", selector: Union[str, _CompoundSelector]) -> List[Any]:
    """Return every proxy in ``document`` that matches ``selector``.

    Accepts either a raw selector string or the AST returned by
    :func:`compile_selector`. Results preserve document order; duplicates
    cannot occur because each candidate appears at most once in the
    enclosing gatherer.
    """
    compound = (
        selector
        if isinstance(selector, _CompoundSelector)
        else compile_selector(selector)
    )
    final_tag = compound.parts[-1][1].tag
    candidates = _GATHERERS[final_tag](document)
    return [c.proxy for c in candidates if _matches_compound(c, compound, document)]


def select_one(
    document: "Document", selector: Union[str, _CompoundSelector]
) -> Optional[Any]:
    """Return the first match for ``selector`` in ``document``, or |None|.

    Equivalent to ``next(iter(select(...)), None)`` but stops gathering
    as soon as a hit is found on the rightmost compound — useful when
    callers only need a single result.
    """
    compound = (
        selector
        if isinstance(selector, _CompoundSelector)
        else compile_selector(selector)
    )
    final_tag = compound.parts[-1][1].tag
    for cand in _GATHERERS[final_tag](document):
        if _matches_compound(cand, compound, document):
            return cand.proxy
    return None
