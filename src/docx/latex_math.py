"""Minimal LaTeX-to-OMML translator for the :mod:`docx.math` proxy layer.

A pragmatic subset — not a full TeX engine — aimed at the common case of
authoring equations in LaTeX and embedding them in Word documents. The
supported grammar is:

* variables — single letters (``x``, ``y``, ``n``, …) → italic ``<m:r>``
* numeric literals — digit runs (``42``, ``3.14``) → upright ``<m:r>``
* binary operators — ``+`` ``-`` ``*`` ``/``
* superscript — ``x^2``, ``x^{i+1}``
* subscript — ``x_i``, ``x_{ij}``
* fractions — ``\\frac{a}{b}``
* square roots — ``\\sqrt{x}``
* parentheses — ``(...)`` → ``<m:d>`` delimiters
* Greek letters — ``\\alpha``, ``\\beta``, …, ``\\Omega``
* equation arrays — ``\\begin{align}... \\\\ ... \\end{align}``

Everything else (matrices, integrals with limits, custom commands, full
LaTeX-to-MathML) is intentionally out of scope. The translator raises
:class:`NotImplementedError` pointing the caller back here when it
encounters unsupported input.

Usage::

    from docx import Document
    from docx.latex_math import latex_to_omml

    doc = Document()
    p = doc.add_paragraph("Euler: ")
    p.add_math_from_latex(r"e^{i \\pi} + 1 = 0")

    # or the standalone form — returns a CT_OMath element ready to paste
    # into any paragraph:
    omath = latex_to_omml(r"\\frac{a+b}{2}")

.. versionadded:: 2026.05.11
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Optional

from ooxml_math import (
    Delimiter,
    EqArray,
    Fraction,
    Lit,
    MathExpr,
    Radical,
    Raw,
    Sub,
    SubSup,
    Sup,
    Text,
    Var,
    oMath,
)

if TYPE_CHECKING:
    from docx.oxml.math import CT_OMath


__all__ = [
    "LatexMathError",
    "latex_to_omml",
]


# ---------------------------------------------------------------------------
# Greek letters (ECMA-376 names → Unicode codepoints)
# ---------------------------------------------------------------------------


# Lower-case Greek
_GREEK_LOWER = {
    "alpha": "α",
    "beta": "β",
    "gamma": "γ",
    "delta": "δ",
    "epsilon": "ϵ",
    "varepsilon": "ε",
    "zeta": "ζ",
    "eta": "η",
    "theta": "θ",
    "vartheta": "ϑ",
    "iota": "ι",
    "kappa": "κ",
    "lambda": "λ",
    "mu": "μ",
    "nu": "ν",
    "xi": "ξ",
    "omicron": "ο",
    "pi": "π",
    "varpi": "ϖ",
    "rho": "ρ",
    "varrho": "ϱ",
    "sigma": "σ",
    "varsigma": "ς",
    "tau": "τ",
    "upsilon": "υ",
    "phi": "ϕ",
    "varphi": "φ",
    "chi": "χ",
    "psi": "ψ",
    "omega": "ω",
}

# Capital Greek
_GREEK_UPPER = {
    "Alpha": "Α",
    "Beta": "Β",
    "Gamma": "Γ",
    "Delta": "Δ",
    "Epsilon": "Ε",
    "Zeta": "Ζ",
    "Eta": "Η",
    "Theta": "Θ",
    "Iota": "Ι",
    "Kappa": "Κ",
    "Lambda": "Λ",
    "Mu": "Μ",
    "Nu": "Ν",
    "Xi": "Ξ",
    "Omicron": "Ο",
    "Pi": "Π",
    "Rho": "Ρ",
    "Sigma": "Σ",
    "Tau": "Τ",
    "Upsilon": "Υ",
    "Phi": "Φ",
    "Chi": "Χ",
    "Psi": "Ψ",
    "Omega": "Ω",
}

_GREEK = {**_GREEK_LOWER, **_GREEK_UPPER}

# Everything the parser is willing to see after a ``\`` — any other
# control word raises NotImplementedError so the caller knows they hit
# unsupported territory.
_KNOWN_COMMANDS = frozenset(
    {"frac", "sqrt", "begin", "end", "\\"} | set(_GREEK.keys())
)

_BINARY_OPS = frozenset({"+", "-", "*", "/"})

_SUPPORTED_ENVIRONMENTS = frozenset({"align", "align*", "aligned"})


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------


class LatexMathError(ValueError):
    """Raised when the LaTeX input is malformed (e.g. unbalanced braces).

    For *unsupported* constructs — valid LaTeX this translator doesn't
    handle — :class:`NotImplementedError` is raised instead, so callers
    can distinguish "your LaTeX is wrong" from "this translator is
    minimal". See the module docstring for the supported subset.

    .. versionadded:: 2026.05.11
    """


# ---------------------------------------------------------------------------
# Tokenizer
# ---------------------------------------------------------------------------


# A token is a tuple ``(kind, value)``. Kinds:
#   "cmd"   — control word (``\frac``, ``\alpha``), value = word without backslash
#   "sep"   — ``\\`` row separator inside ``\begin{align}...\end{align}``
#   "char"  — any other single printable character (letter, digit, operator,
#             brace, paren, ``&``)
#   "eof"   — sentinel used internally


_Token = tuple  # just tuple[str, str] — written as tuple for py3.9 forward compat


def _tokenize(src: str) -> "List[_Token]":
    """Return the token stream for *src*.

    Whitespace is dropped (LaTeX is whitespace-insensitive inside an
    equation body). Raises :class:`LatexMathError` on unterminated
    control words at end-of-input.
    """
    tokens: List[_Token] = []
    i = 0
    n = len(src)
    while i < n:
        c = src[i]
        if c.isspace():
            i += 1
            continue
        if c == "\\":
            # ``\\`` — row separator
            if i + 1 < n and src[i + 1] == "\\":
                tokens.append(("sep", "\\\\"))
                i += 2
                continue
            # Control word: ``\`` + one-or-more letters.
            j = i + 1
            while j < n and src[j].isalpha():
                j += 1
            if j > i + 1:
                tokens.append(("cmd", src[i + 1 : j]))
                i = j
                continue
            # Control symbol: ``\`` + single non-alpha char (``\,``, ``\!``,
            # ``\{``, ``\;``, …). We don't support any of these, but
            # preserve them as ``cmd`` tokens so the parser can raise
            # NotImplementedError with a helpful message instead of
            # LatexMathError (which would be misleading — the input is
            # well-formed LaTeX, just outside our subset).
            if j < n:
                tokens.append(("cmd", src[j]))
                i = j + 1
                continue
            raise LatexMathError(
                f"lone backslash at end of input (position {i})"
            )
        # Every other character is a single-char token.
        tokens.append(("char", c))
        i += 1
    return tokens


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------


class _Parser:
    """Recursive-descent parser over the token stream.

    The parser is deliberately small — it builds a flat list of
    :class:`MathExpr` children for an ``oMath`` root. Precedence is
    flat too: ``+-*/`` are emitted as plain operator runs, so ``a+b*c``
    is laid out left-to-right without OMML-level grouping (matching the
    behaviour of Word when you type the same expression in its equation
    editor).
    """

    def __init__(self, tokens: "List[_Token]") -> None:
        self._tokens = tokens
        self._pos = 0

    # -- helpers -----------------------------------------------------------

    def _peek(self, offset: int = 0) -> "Optional[_Token]":
        idx = self._pos + offset
        if idx >= len(self._tokens):
            return None
        return self._tokens[idx]

    def _advance(self) -> "Optional[_Token]":
        tok = self._peek()
        if tok is not None:
            self._pos += 1
        return tok

    def _expect_char(self, ch: str) -> None:
        tok = self._peek()
        if tok is None or tok != ("char", ch):
            got = "EOF" if tok is None else f"{tok[0]}({tok[1]!r})"
            raise LatexMathError(f"expected {ch!r}, got {got}")
        self._advance()

    # -- top level ---------------------------------------------------------

    def parse(self) -> "List[MathExpr]":
        """Consume the whole token stream as a flat expression list."""
        # Top-level ``\begin{align}...\end{align}`` is handled here so
        # the caller sees an :class:`EqArray`-containing expression
        # rather than a flat sequence with stray separators.
        if self._is_begin_env():
            env_name = self._read_environment_header()
            if env_name not in _SUPPORTED_ENVIRONMENTS:
                raise NotImplementedError(
                    f"environment {env_name!r} is not supported by "
                    "docx.latex_math. Supported: align, align*, aligned. "
                    "See the module docstring for the full supported subset."
                )
            array = self._parse_equation_array_body()
            self._read_environment_footer(env_name)
            # Nothing else allowed after a full align environment.
            if self._peek() is not None:
                raise LatexMathError(
                    "content after \\end{align} is not supported"
                )
            return [array]
        children = self._parse_sequence(stop_chars=set())
        return children

    # -- environment helpers ----------------------------------------------

    def _is_begin_env(self) -> bool:
        tok = self._peek()
        return tok is not None and tok[0] == "cmd" and tok[1] == "begin"

    def _read_environment_header(self) -> str:
        # Consume ``\begin`` + ``{name}``.
        self._advance()  # \begin
        self._expect_char("{")
        name_chars: List[str] = []
        while True:
            tok = self._peek()
            if tok is None:
                raise LatexMathError(
                    "unterminated \\begin{...} — missing closing brace"
                )
            if tok == ("char", "}"):
                self._advance()
                break
            if tok[0] != "char":
                raise LatexMathError(
                    f"invalid environment name token {tok!r}"
                )
            name_chars.append(tok[1])
            self._advance()
        return "".join(name_chars)

    def _read_environment_footer(self, expected: str) -> None:
        tok = self._peek()
        if tok is None or tok != ("cmd", "end"):
            raise LatexMathError(
                f"expected \\end{{{expected}}} at end of environment"
            )
        self._advance()  # \end
        self._expect_char("{")
        name_chars: List[str] = []
        while True:
            t = self._peek()
            if t is None:
                raise LatexMathError(
                    "unterminated \\end{...} — missing closing brace"
                )
            if t == ("char", "}"):
                self._advance()
                break
            if t[0] != "char":
                raise LatexMathError(
                    f"invalid environment name token {t!r}"
                )
            name_chars.append(t[1])
            self._advance()
        got = "".join(name_chars)
        if got != expected:
            raise LatexMathError(
                f"mismatched environment — \\begin{{{expected}}} vs "
                f"\\end{{{got}}}"
            )

    def _parse_equation_array_body(self) -> MathExpr:
        """Consume tokens until an ``\\end`` command, splitting on ``\\\\``."""
        rows: List[MathExpr] = []
        current: List[MathExpr] = []
        while True:
            tok = self._peek()
            if tok is None:
                raise LatexMathError(
                    "unterminated equation-array body (missing \\end)"
                )
            if tok[0] == "cmd" and tok[1] == "end":
                break
            if tok[0] == "sep":
                rows.append(_wrap_row(current))
                current = []
                self._advance()
                continue
            current.append(self._parse_atom_with_scripts())
        rows.append(_wrap_row(current))
        return EqArray(rows)

    # -- core parsing ------------------------------------------------------

    def _parse_sequence(
        self, stop_chars: "set[str]"
    ) -> "List[MathExpr]":
        """Read atoms until EOF or one of *stop_chars* is seen."""
        out: List[MathExpr] = []
        while True:
            tok = self._peek()
            if tok is None:
                break
            if tok[0] == "char" and tok[1] in stop_chars:
                break
            if tok[0] == "sep":
                raise LatexMathError(
                    "row separator \\\\ is only valid inside a supported "
                    "environment (\\begin{align}...\\end{align})"
                )
            out.append(self._parse_atom_with_scripts())
        return out

    def _parse_atom_with_scripts(self) -> MathExpr:
        """Parse one atom, then optionally ``_`` / ``^`` suffixes."""
        base = self._parse_atom()
        sub: "Optional[MathExpr]" = None
        sup: "Optional[MathExpr]" = None
        while True:
            tok = self._peek()
            if tok == ("char", "_"):
                if sub is not None:
                    raise LatexMathError("duplicate '_' on the same atom")
                self._advance()
                sub = self._parse_group_or_atom()
                continue
            if tok == ("char", "^"):
                if sup is not None:
                    raise LatexMathError("duplicate '^' on the same atom")
                self._advance()
                sup = self._parse_group_or_atom()
                continue
            break
        if sub is not None and sup is not None:
            return SubSup(base, sub, sup)
        if sub is not None:
            return Sub(base, sub)
        if sup is not None:
            return Sup(base, sup)
        return base

    def _parse_atom(self) -> MathExpr:
        tok = self._peek()
        if tok is None:
            raise LatexMathError("unexpected end of input — expected an atom")
        kind, value = tok
        if kind == "cmd":
            return self._parse_command()
        assert kind == "char"
        if value == "{":
            return self._parse_group_as_expr()
        if value == "(":
            return self._parse_parenthesised()
        if value in _BINARY_OPS:
            self._advance()
            return Text(value, italic=False)
        if value == "=":
            self._advance()
            return Text("=", italic=False)
        if value.isdigit() or value == ".":
            return self._parse_number()
        if value.isalpha():
            self._advance()
            return Var(value)
        # Any other raw character (``,``, ``!``, ``|`` …) — unsupported
        # at this level of the translator.
        raise NotImplementedError(
            f"character {value!r} is not supported by docx.latex_math. "
            "Supported atoms: letters, digits, + - * / = ( ) { } _ ^, "
            "\\frac, \\sqrt, \\begin{align}...\\end{align}, Greek commands. "
            "See the module docstring for the full supported subset."
        )

    def _parse_parenthesised(self) -> MathExpr:
        """Consume ``(...)`` and return a :class:`Delimiter` around the body."""
        self._expect_char("(")
        children = self._parse_sequence(stop_chars={")"})
        self._expect_char(")")
        # A Delimiter takes one or more arguments; ``(a+b)`` is a single
        # argument (a+b), not two comma-separated args, so we pack the
        # whole run into a single :class:`MathExprLike` using our group
        # helper.
        if len(children) == 0:
            return Delimiter(Var(""), begin="(", end=")")
        if len(children) == 1:
            return Delimiter(children[0], begin="(", end=")")
        return Delimiter(_Group(children), begin="(", end=")")

    def _parse_number(self) -> MathExpr:
        """Read a run of digits (with at most one ``.``) as a single Lit."""
        chars: List[str] = []
        saw_dot = False
        while True:
            tok = self._peek()
            if tok is None or tok[0] != "char":
                break
            ch = tok[1]
            if ch.isdigit():
                chars.append(ch)
                self._advance()
                continue
            if ch == "." and not saw_dot:
                # Peek ahead — ``.`` is only part of the number when
                # followed by another digit (otherwise it's a standalone
                # punctuator, which we reject above anyway).
                nxt = self._peek(1)
                if nxt is not None and nxt[0] == "char" and nxt[1].isdigit():
                    chars.append(ch)
                    saw_dot = True
                    self._advance()
                    continue
            break
        return Lit("".join(chars))

    def _parse_command(self) -> MathExpr:
        tok = self._advance()
        assert tok is not None and tok[0] == "cmd"
        name = tok[1]
        if name == "frac":
            num = self._parse_required_group()
            den = self._parse_required_group()
            return Fraction(num, den)
        if name == "sqrt":
            radicand = self._parse_required_group()
            return Radical(radicand)
        if name in _GREEK:
            # Greek letters typeset italic by default (Word math-font
            # behaviour), so emit as a plain Var holding the Unicode
            # codepoint.
            return Var(_GREEK[name])
        if name in ("begin", "end"):
            raise LatexMathError(
                f"bare \\{name} is only valid inside an environment header"
            )
        if name not in _KNOWN_COMMANDS:
            raise NotImplementedError(
                f"LaTeX command \\{name} is not supported by "
                "docx.latex_math. Supported commands: \\frac, \\sqrt, "
                "\\begin{align}...\\end{align}, and the common Greek "
                "letters (\\alpha ... \\omega, \\Gamma ... \\Omega). "
                "See the module docstring for the full supported subset."
            )
        # Reachable only if a name is in _KNOWN_COMMANDS but not handled
        # above — a programmer error in this file.
        raise AssertionError(f"unhandled known command: \\{name}")  # pragma: no cover

    def _parse_group_or_atom(self) -> MathExpr:
        """Parse either a ``{...}`` group or a single atom (for ``^``/``_``)."""
        tok = self._peek()
        if tok == ("char", "{"):
            return self._parse_group_as_expr()
        return self._parse_atom()

    def _parse_required_group(self) -> MathExpr:
        """Consume ``{...}`` and return its contents as one :class:`MathExpr`."""
        tok = self._peek()
        if tok != ("char", "{"):
            raise LatexMathError(
                "expected a brace-delimited group '{...}'"
            )
        return self._parse_group_as_expr()

    def _parse_group_as_expr(self) -> MathExpr:
        """Consume ``{...}`` and collapse its children into a single expr."""
        self._expect_char("{")
        children = self._parse_sequence(stop_chars={"}"})
        self._expect_char("}")
        if len(children) == 0:
            # Empty group — render as an empty variable (Word tolerates
            # an empty <m:r>).
            return Var("")
        if len(children) == 1:
            return children[0]
        # Multiple children → wrap in an ``oMath`` (it's the only generic
        # container that's a MathExpr). Inside a fraction / sqrt arg the
        # proxy layer strips the ``<m:oMath>`` envelope? No — _build_arg
        # nests whatever element it's given. A bare ``<m:oMath>`` inside
        # ``<m:num>`` is malformed. Use a degenerate delimiter instead:
        # ``<m:d>`` with no opener/closer and a single ``<m:e>`` argument
        # grouping the children. To keep the rendering clean, we emit
        # an invisible delimiter pair (``"〈"``?) — simpler: just
        # concatenate via a :class:`_Group` helper that wraps children
        # under a fresh ``<m:e>``-compatible container using ``Delimiter``
        # with empty begin/end chars.
        return _Group(children)


# ---------------------------------------------------------------------------
# Helpers used by the parser
# ---------------------------------------------------------------------------


def _Group(children: "List[MathExpr]") -> MathExpr:
    """Return a :class:`MathExpr` that renders *children* as a flat sequence.

    OMML argument sites (``<m:num>``, ``<m:den>``, ``<m:e>``, …) already
    accept a mixed-content run of children, so the proxy-layer
    :func:`_build_arg` helper appends whatever you hand it. But
    :class:`Fraction` / :class:`Radical` / :class:`Sub` / ``Sup`` all
    take a *single* ``MathExprLike`` per slot. To pack ``{a+b}`` into
    one slot we need a single element whose children are ``a``, ``+``,
    ``b``.

    We use :class:`Delimiter` with empty ``begin``/``end`` — the OMML
    schema allows empty ``<m:begChr>`` / ``<m:endChr>`` and Word
    renders it as an invisible grouping, which is exactly what LaTeX
    ``{...}`` means.
    """
    # _build_arg auto-wraps each expr in its own ``<m:e>`` — which is
    # the right semantic for delimiter arguments but the wrong one for
    # a flat group. So instead of a Delimiter we hand-build a single
    # ``<m:d>`` with one ``<m:e>`` containing every child in order.
    from ooxml_math.constants import NS_M
    from ooxml_math.oxml import _OxmlElement

    M = "{%s}" % NS_M
    d = _OxmlElement("m:d")
    # Empty begin / end chars render the delimiter invisible (ECMA-376
    # §22.1.2.17 m:dPr/m:begChr default is ``(`` but an empty-value
    # attribute suppresses the glyph, matching the way Word writes
    # LaTeX-style invisible groupings).
    dPr = _OxmlElement("m:dPr")
    begChr = _OxmlElement("m:begChr")
    begChr.set(M + "val", "")
    endChr = _OxmlElement("m:endChr")
    endChr.set(M + "val", "")
    dPr.append(begChr)
    dPr.append(endChr)
    d.append(dPr)
    e = _OxmlElement("m:e")
    for child in children:
        # pyright: reportPrivateUsage=false
        e.append(child._element)
    d.append(e)
    return Raw(d)


def _wrap_row(children: "List[MathExpr]") -> MathExpr:
    """Return a single :class:`MathExpr` holding a row of *children*."""
    if len(children) == 0:
        return Var("")
    if len(children) == 1:
        return children[0]
    return _Group(children)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def latex_to_omml(latex: str) -> "CT_OMath":
    """Translate a LaTeX math body into an OMML ``<m:oMath>`` element.

    *latex* is the body only — ``$$…$$`` / ``\\[…\\]`` / ``\\(…\\)``
    wrappers are **not** part of the input. See the module docstring for
    the full supported subset.

    Returns the underlying :class:`~docx.oxml.math.CT_OMath` element
    (the same type you'd get out of :attr:`Equation.xml_element`).
    Parents are unattached — you typically hand the result to
    :meth:`~docx.text.paragraph.Paragraph.add_equation` or
    :meth:`~docx.text.paragraph.Paragraph.add_math_from_latex`.

    Raises :class:`LatexMathError` on malformed input; raises
    :class:`NotImplementedError` on well-formed LaTeX that uses a
    construct outside the supported subset.

    .. versionadded:: 2026.05.11
    """
    if not isinstance(latex, str):
        raise TypeError(
            f"latex must be a str, got {type(latex).__name__}"
        )
    tokens = _tokenize(latex)
    parser = _Parser(tokens)
    children = parser.parse()
    root = oMath(*children)
    # Cast to CT_OMath — oMath._element is built via _OxmlElement("m:oMath")
    # which goes through the ooxml-xmlchemy parser registry and returns a
    # CT_OMath instance.
    from docx.oxml.math import CT_OMath as _CT_OMath

    element = root._element  # noqa: SLF001 — proxy-layer contract
    assert isinstance(element, _CT_OMath), (
        f"internal error: oMath proxy wrapped a {type(element).__name__} "
        "instead of a CT_OMath"
    )
    return element
