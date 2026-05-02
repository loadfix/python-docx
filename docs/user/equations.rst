.. _equations:

Working with equations
======================

Word stores mathematical expressions as *Office Math* (OMML, the ``m:`` namespace)
rather than as text runs. An equation lives in one of two container elements:

- ``<m:oMath>`` — an inline equation, embedded in a run-level position
  inside a paragraph, flowing with the surrounding text;
- ``<m:oMathPara>`` — a *display-mode* equation, which Word centers on its
  own line and renders at a larger size. ``m:oMathPara`` always wraps
  exactly one ``m:oMath`` element and carries its own formatting in
  ``m:oMathParaPr``.

*python-docx* provides a read-only |Equation| proxy over either element, plus a
small family of *builder* functions that emit OMML XML strings for the most
common single-node idioms (identifiers, fractions, sub/superscripts, radicals).
Import/export for LaTeX or MathML is intentionally out of scope — the OMML XML
string is the exchange format.


Reading equations from a document
---------------------------------

The document-level :attr:`.Document.equations` property returns every top-level
equation found by walking the document body::

    >>> from docx import Document
    >>> document = Document("has-equations.docx")
    >>> document.equations
    [<docx.equations.Equation object at 0x02468ACE>, ...]
    >>> len(document.equations)
    2

Each entry is an |Equation| proxy. The walk yields ``m:oMathPara`` wrappers
whole; an inline ``m:oMath`` nested inside a ``m:oMathPara`` is represented
once, by the enclosing wrapper, not as two separate equations. Equations inside
headers, footers, footnotes, endnotes, and comments are **not** included in
this collection — they belong to the corresponding story container, and are
accessible via :attr:`.Paragraph.equations` on paragraphs within those stories.

Paragraph-level access mirrors the document-level shape::

    >>> paragraph = document.paragraphs[1]
    >>> [e.text for e in paragraph.equations]
    ['x']

.. note::

   The paragraph and document equation walks are *read-only iterators*. They
   reflect the current XML tree; they do not expose add/remove operations
   directly. Creating an equation is always done by appending OMML XML via
   :meth:`.Paragraph.add_equation` (see below).


Inspecting an equation
----------------------

The |Equation| proxy exposes three read-only properties over the wrapped
``m:oMath`` / ``m:oMathPara`` element:

* :attr:`~.Equation.text` — a best-effort, *flattened* plain-text rendering
  that concatenates every descendant ``m:t`` element's text. Structure
  (fractions, sub/superscripts, radicals) is stripped, which is usually good
  enough for search indexing or quick previews but loses the mathematical
  meaning. Use :attr:`~.Equation.raw_xml` when fidelity matters.
* :attr:`~.Equation.raw_xml` — the serialized OMML XML for this equation, as
  UTF-8 bytes, with all namespace declarations preserved. Callers who want to
  reason about the tree should hand these bytes to their own XML parser.
* :attr:`~.Equation.is_display_mode` — |True| when the wrapped element is
  ``m:oMathPara``, |False| when it is a bare inline ``m:oMath``::

    >>> equation = document.equations[0]
    >>> equation.text
    'x'
    >>> equation.raw_xml[:48]
    b'<m:oMath xmlns:m="http://schemas.openxmlfor'
    >>> equation.is_display_mode
    False

The underlying lxml element is also exposed via
:attr:`~.Equation.xml_element` for advanced callers who need direct XPath
access into the tree.

If you have an OMML XML string produced out of band (from a template, an
XSLT pass, another tool) you can wrap it directly::

    >>> from docx.equations import Equation
    >>> xml = '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>e</m:t></m:r></m:oMath>'
    >>> equation = Equation.from_omml_xml(xml)
    >>> equation.text
    'e'

:meth:`.Equation.from_omml_xml` raises :class:`ValueError` when the root
element is neither ``m:oMath`` nor ``m:oMathPara``. Namespace declarations for
the ``m:`` prefix must be present on the root element (or an ancestor); the
caller is responsible for including them.


Appending an equation to a paragraph
------------------------------------

:meth:`.Paragraph.add_equation` parses an OMML XML string and appends the
resulting element to the paragraph, returning the wrapping |Equation|::

    >>> paragraph = document.add_paragraph("The variable ")
    >>> xml = '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath>'
    >>> equation = paragraph.add_equation(xml)
    >>> equation.text
    'x'
    >>> equation.is_display_mode
    False

Passing ``display_mode=True`` wraps a bare ``m:oMath`` in an ``m:oMathPara``
before appending, turning it into a centered display equation. If the supplied
XML is already an ``m:oMathPara``, it is appended unchanged regardless of the
flag::

    >>> equation = paragraph.add_equation(xml, display_mode=True)
    >>> equation.is_display_mode
    True


Builder helpers
---------------

Hand-authoring OMML XML is possible but verbose. The
:mod:`docx.equations` module ships a small family of builder functions that
each return a complete, parseable ``m:oMath`` fragment with the namespace
declaration already in place. Their output is suitable for passing directly to
:meth:`.Paragraph.add_equation` or :meth:`.Equation.from_omml_xml`::

    >>> from docx.equations import (
    ...     build_identifier, build_fraction,
    ...     build_superscript, build_subscript, build_radical,
    ... )

These helpers cover the everyday shapes. When you need nested structure —
a fraction whose numerator is itself a superscript, for instance — hand-author
the OMML or compose the builders' output with an XML tree library of your
choice.


build_identifier -- plain identifiers and literals
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

:func:`~docx.equations.build_identifier` wraps a short text span in a single
``<m:r><m:t>…</m:t></m:r>`` run inside an ``m:oMath`` element. It is the
correct building block for a single-letter variable, a Greek letter, or a
short keyword::

    >>> build_identifier("x")
    '<m:oMath xmlns:m="...">...<m:r><m:t>x</m:t></m:r></m:oMath>'
    >>> build_identifier("χ")  # Greek chi
    '<m:oMath xmlns:m="...">...<m:r><m:t>χ</m:t></m:r></m:oMath>'

The `text` argument is XML-escaped, so identifiers that happen to contain
characters like ``<`` or ``&`` are safe::

    >>> build_identifier("a<b")
    '<m:oMath xmlns:m="...">...<m:r><m:t>a&lt;b</m:t></m:r></m:oMath>'

Word does not italicize the glyph based on the identifier's content; styling
is controlled by the ``<m:rPr>`` run-property child, which this builder does
not emit. The default Word rendering for ``<m:t>`` inside a run with no
explicit ``<m:rPr><m:sty m:val="p"/></m:rPr>`` is italic, matching convention
for mathematical variables.


build_fraction -- stacked fractions with a horizontal bar
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

:func:`~docx.equations.build_fraction` emits an ``m:f`` element (a "fraction")
containing an ``m:num`` (numerator) and ``m:den`` (denominator), each wrapped
around a single run. The ``m:fPr`` child carries ``<m:type m:val="bar"/>`` to
select the stacked horizontal-bar appearance (``"bar"`` is the default; other
types are ``"lin"`` for linear *a/b*, ``"noBar"`` for stacked without a bar,
and ``"skw"`` for skewed)::

    >>> build_fraction("a", "b")
    '<m:oMath xmlns:m="..."><m:f><m:fPr><m:type m:val="bar"/></m:fPr>
     <m:num><m:r><m:t>a</m:t></m:r></m:num>
     <m:den><m:r><m:t>b</m:t></m:r></m:den>
     </m:f></m:oMath>'

Both arguments are wrapped as a single ``m:r``/``m:t`` run — the builder does
not parse its inputs. To nest a fraction inside a fraction, or to place a
superscript in the numerator, hand-author the OMML around the builder's
output.

The flattened :attr:`.Equation.text` of a fraction concatenates the numerator
and denominator text in that order::

    >>> equation = Equation.from_omml_xml(build_fraction("a", "b"))
    >>> equation.text
    'ab'


build_superscript -- exponents
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

:func:`~docx.equations.build_superscript` emits an ``m:sSup`` element (a
"script-super") with ``m:e`` (the base) and ``m:sup`` (the exponent) children.
Each is wrapped around a single run::

    >>> build_superscript("x", "2")
    '<m:oMath xmlns:m="..."><m:sSup>
     <m:e><m:r><m:t>x</m:t></m:r></m:e>
     <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
     </m:sSup></m:oMath>'


build_subscript -- subscripts
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

:func:`~docx.equations.build_subscript` is the mirror image of
:func:`~docx.equations.build_superscript`: it emits an ``m:sSub`` element
("script-sub") with ``m:e`` (the base) and ``m:sub`` (the subscript)::

    >>> build_subscript("x", "i")
    '<m:oMath xmlns:m="..."><m:sSub>
     <m:e><m:r><m:t>x</m:t></m:r></m:e>
     <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
     </m:sSub></m:oMath>'

For an identifier that carries both a subscript *and* a superscript
simultaneously (``x_i^2``), hand-author an ``m:sSubSup`` element — the
builders do not cover that case.


build_radical -- square roots and nth roots
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

:func:`~docx.equations.build_radical` emits an ``m:rad`` element (a "radical")
with an optional ``m:deg`` (degree) and a required ``m:e`` (the radicand).
When `degree_text` is |None| (the default) an empty ``<m:deg/>`` is written,
which Word renders as a square-root glyph::

    >>> build_radical("x")
    '<m:oMath xmlns:m="..."><m:rad><m:deg/>
     <m:e><m:r><m:t>x</m:t></m:r></m:e>
     </m:rad></m:oMath>'
    >>> build_radical("x", "3")  # cube root
    '<m:oMath xmlns:m="..."><m:rad>
     <m:deg><m:r><m:t>3</m:t></m:r></m:deg>
     <m:e><m:r><m:t>x</m:t></m:r></m:e>
     </m:rad></m:oMath>'


Composing builders with add_equation
------------------------------------

The intended workflow is to build an OMML fragment with one of the helpers,
then hand it to :meth:`.Paragraph.add_equation`. The example below composes
a sentence with an inline fraction equation::

    >>> from docx import Document
    >>> from docx.equations import build_fraction
    >>> document = Document()
    >>> paragraph = document.add_paragraph("The ratio is ")
    >>> equation = paragraph.add_equation(build_fraction("a", "b"))
    >>> equation.text
    'ab'
    >>> document.save("ratio.docx")


Limitations
-----------

The builder helpers are intentionally minimal. They each return a single
top-level mathematical node wrapped in ``m:oMath``, suitable for the most
common idioms but not for nesting. In particular:

- Each builder's argument becomes a single ``<m:t>`` text run. Numerators,
  denominators, bases, exponents, degrees, and radicands cannot themselves
  contain further builder output directly — the ``str`` arguments are not
  parsed. To nest (a fraction whose numerator is a superscript, for instance),
  hand-author the OMML.
- No run properties (``<m:rPr>``) are emitted. Style (italic/roman, color,
  size) is inherited from the paragraph.
- Combined sub-and-super (``m:sSubSup``), matrices (``m:m``), delimiters
  (``m:d``), and accent marks (``m:acc``) are not covered by a builder. Read
  access via |Equation| still works for these elements — their text content
  is flattened into :attr:`~.Equation.text` — but creating them requires
  hand-authored OMML.
- LaTeX or MathML input/output is out of scope. The OMML XML string remains
  the authoritative exchange format.
