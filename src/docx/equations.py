"""Equation-related proxy types and a minimal OMML builder helper module.

An :class:`Equation` wraps an ``<m:oMath>`` or ``<m:oMathPara>`` element. Reading
is the primary use case: the raw OMML XML is exposed as bytes, along with a
best-effort plain-text rendering and the display-mode flag.

The module also exposes a tiny set of *builder* functions that return OMML XML
strings for common idioms — fractions, sub/superscripts, radicals, identifiers.
These are convenience helpers; callers who want fuller fidelity should hand-
author OMML and pass it to :meth:`Equation.from_omml_xml`.

LaTeX/MathML import/export is intentionally out of scope (see GitHub issue
#113). The OMML XML string remains the authoritative exchange format here.
"""

from __future__ import annotations

from typing import TYPE_CHECKING
from xml.sax.saxutils import escape as _xml_escape

from lxml import etree

from docx.oxml.math import CT_OMath, CT_OMathPara
from docx.oxml.ns import nsmap, qn
from docx.oxml.parser import parse_xml

if TYPE_CHECKING:
    from docx.oxml.xmlchemy import BaseOxmlElement


_M_NS = nsmap["m"]
_M_OMATH = qn("m:oMath")
_M_OMATH_PARA = qn("m:oMathPara")


class Equation:
    """Proxy for an Office Math (OMML) expression.

    Wraps either a top-level ``m:oMath`` element (inline) or an
    ``m:oMathPara`` element (display-mode paragraph). The wrapped element is
    accessible via :attr:`xml_element` for advanced use cases.
    """

    def __init__(self, element: CT_OMath | CT_OMathPara):
        if element.tag not in (_M_OMATH, _M_OMATH_PARA):
            raise ValueError(
                "Equation must wrap an m:oMath or m:oMathPara element, got %r"
                % element.tag
            )
        self._element = element

    @property
    def xml_element(self) -> CT_OMath | CT_OMathPara:
        """The raw lxml element (``m:oMath`` or ``m:oMathPara``)."""
        return self._element

    @property
    def raw_xml(self) -> bytes:
        """The OMML XML for this equation, as bytes.

        Namespaces are preserved; callers can hand this to their own XML
        parser for deeper inspection.
        """
        return etree.tostring(self._element, encoding="utf-8")

    @property
    def text(self) -> str:
        """A best-effort plain-text rendering of the equation.

        Concatenates the text of every descendant ``<m:t>`` element. Structure
        (fractions, superscripts, radicals, …) is flattened, which is usually
        good enough for search or preview display.
        """
        return "".join(t.text or "" for t in self._element.xpath(".//m:t"))

    @property
    def is_display_mode(self) -> bool:
        """|True| when the equation is wrapped in ``m:oMathPara`` (display mode)."""
        return self._element.tag == _M_OMATH_PARA

    @classmethod
    def from_omml_xml(cls, xml_string: str | bytes) -> Equation:
        """Return a new |Equation| parsed from an OMML XML string.

        `xml_string` must be a well-formed XML document whose root element is
        either ``m:oMath`` or ``m:oMathPara``. Namespace declarations for the
        ``m`` prefix must be present on the root element (or an ancestor); the
        caller is responsible for including them.

        Raises :class:`ValueError` when the root element has a different tag.
        """
        if isinstance(xml_string, str):
            xml_string = xml_string.encode("utf-8")
        element = parse_xml(xml_string)
        if element.tag not in (_M_OMATH, _M_OMATH_PARA):
            raise ValueError(
                "OMML root must be m:oMath or m:oMathPara; got %r" % element.tag
            )
        return cls(element)  # pyright: ignore[reportArgumentType]


# ---------------------------------------------------------------------------
# Internal helpers for EquationBuilder
# ---------------------------------------------------------------------------


def _omath_open() -> str:
    return '<m:oMath xmlns:m="%s">' % _M_NS


def _run(text: str) -> str:
    """Return ``<m:r><m:t>text</m:t></m:r>`` with XML-escaped text."""
    return "<m:r><m:t>%s</m:t></m:r>" % _xml_escape(text)


# ---------------------------------------------------------------------------
# EquationBuilder — small factory functions returning OMML XML strings.
# Each return value is a complete, parseable ``m:oMath`` fragment suitable for
# passing to :meth:`Equation.from_omml_xml` or ``paragraph.add_equation``.
# ---------------------------------------------------------------------------


def build_identifier(text: str) -> str:
    """Return an ``m:oMath`` expressing a plain identifier or literal.

    `text` is XML-escaped. Produces::

        <m:oMath><m:r><m:t>text</m:t></m:r></m:oMath>
    """
    return "%s%s</m:oMath>" % (_omath_open(), _run(text))


def build_fraction(numerator_text: str, denominator_text: str) -> str:
    """Return an ``m:oMath`` expressing ``numerator_text / denominator_text``.

    Produces a stacked fraction with a horizontal bar (``m:type=bar``). Both
    arguments are wrapped as a single ``m:r``/``m:t`` run — use
    :meth:`Equation.from_omml_xml` directly if you need nested structure.
    """
    return (
        "%s<m:f>"
        '<m:fPr><m:type m:val="bar"/></m:fPr>'
        "<m:num>%s</m:num>"
        "<m:den>%s</m:den>"
        "</m:f></m:oMath>"
    ) % (_omath_open(), _run(numerator_text), _run(denominator_text))


def build_superscript(base_text: str, exponent_text: str) -> str:
    """Return an ``m:oMath`` expressing ``base_text`` raised to ``exponent_text``.

    Uses the ``m:sSup`` element.
    """
    return (
        "%s<m:sSup>"
        "<m:e>%s</m:e>"
        "<m:sup>%s</m:sup>"
        "</m:sSup></m:oMath>"
    ) % (_omath_open(), _run(base_text), _run(exponent_text))


def build_subscript(base_text: str, subscript_text: str) -> str:
    """Return an ``m:oMath`` expressing ``base_text`` with a subscript.

    Uses the ``m:sSub`` element.
    """
    return (
        "%s<m:sSub>"
        "<m:e>%s</m:e>"
        "<m:sub>%s</m:sub>"
        "</m:sSub></m:oMath>"
    ) % (_omath_open(), _run(base_text), _run(subscript_text))


def build_radical(expr_text: str, degree_text: str | None = None) -> str:
    """Return an ``m:oMath`` expressing a radical (√ by default, nth-root when given).

    When `degree_text` is |None|, a square-root (no degree) is emitted. When
    given, a degree run is added inside ``m:deg`` to produce e.g. ∛ for a
    degree of ``"3"``.
    """
    deg_xml = "<m:deg>%s</m:deg>" % _run(degree_text) if degree_text else "<m:deg/>"
    return (
        "%s<m:rad>"
        "%s"
        "<m:e>%s</m:e>"
        "</m:rad></m:oMath>"
    ) % (_omath_open(), deg_xml, _run(expr_text))


# ---------------------------------------------------------------------------
# Paragraph helper — constructs an equation element from an OMML XML string
# and (optionally) wraps it in ``m:oMathPara`` for display-mode.
# ---------------------------------------------------------------------------


def _make_equation_element(
    omml_xml: str | bytes, display_mode: bool = False
) -> BaseOxmlElement:
    """Parse `omml_xml` and return an element ready to append to a paragraph.

    When `display_mode` is |True| and the parsed root is bare ``m:oMath``,
    it is wrapped in ``m:oMathPara``. When the root is already ``m:oMathPara``,
    it is returned unchanged regardless of `display_mode`.
    """
    if isinstance(omml_xml, str):
        omml_xml = omml_xml.encode("utf-8")
    element = parse_xml(omml_xml)
    if element.tag not in (_M_OMATH, _M_OMATH_PARA):
        raise ValueError(
            "OMML root must be m:oMath or m:oMathPara; got %r" % element.tag
        )
    if display_mode and element.tag == _M_OMATH:
        # -- build a wrapper <m:oMathPara> with empty <m:oMathParaPr> --
        wrapper_xml = (
            '<m:oMathPara xmlns:m="%s"><m:oMathParaPr/></m:oMathPara>' % _M_NS
        )
        wrapper = parse_xml(wrapper_xml.encode("utf-8"))
        wrapper.append(element)
        return wrapper
    return element
