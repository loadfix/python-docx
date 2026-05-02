
API basics
==========

The API for |docx| is designed to make doing simple things simple, while
allowing more complex results to be achieved with a modest and incremental
investment of understanding.

It's possible to create a basic document using only a single object, the
|api-Document| object returned when opening a file. The methods on
|api-Document| allow *block-level* objects to be added to the end of the
document. Block-level objects include paragraphs, inline pictures, and tables.
Headings, bullets, and numbered lists are simply paragraphs with a particular
style applied.

In this way, a document can be "written" from top to bottom, roughly like
a person would if they knew exactly what they wanted to say This basic use
case, where content is always added to the end of the document, is expected to
account for perhaps 80% of actual use cases, so it's a priority to make it as
simple as possible without compromising the power of the overall API.


Inline objects
--------------

Each block-level method on |api-Document|, such as ``add_paragraph()``, returns
the block-level object created. Often the reference is unneeded; but when
inline objects must be created individually, you'll need the block-item
reference to do it.

... add example here as API solidifies ...


Architecture: proxy, part, oxml
-------------------------------

Most |docx| users never need to look below the friendly ``Document``,
``Paragraph``, ``Run``, and ``Table`` APIs. That's by design. When a question
gets more interesting however — "how is this actually stored on disk?",
"why does my edit round-trip that way?", or "how can I reach a feature the
high-level API doesn't yet cover?" — it helps to understand the three layers
the library is built on.

|docx| is organized as a stack::

    Document API   (src/docx/document.py, src/docx/text/*.py, ...)
        |  Proxy objects wrapping oxml elements
    Parts Layer    (src/docx/parts/*.py)
        |  XmlPart subclasses owning XML trees and relationships
    oxml Layer     (src/docx/oxml/*.py)
        |  CT_* element classes extending lxml.etree.ElementBase
    lxml           (XML parsing and serialization)

Each layer is narrow and does exactly one job:

* **Document API (proxy).** The classes you import from |docx| — ``Document``,
  ``Paragraph``, ``Run``, ``Table``, ``Section``, ``Footnote`` — are *proxy
  objects*. They hold no content of their own; they wrap a single oxml element
  and expose an ergonomic, Pythonic interface over it. Proxies inherit from
  ``ElementProxy``, ``StoryChild``, or ``BlockItemContainer`` depending on
  what they wrap.

* **Parts Layer.** A `.docx` file is really a ZIP package (OPC) containing
  several XML *parts* — the main document, styles, numbering, comments,
  footnotes, and so on — joined together by relationships. Each part is
  represented by a subclass of ``XmlPart`` (or ``StoryPart``), which owns the
  parsed XML tree for that part and knows how to find related parts. The
  parts layer is where lazy creation happens: a footnotes part, for example,
  is only attached to the document when something first asks for it.

* **oxml Layer.** The ``CT_*`` classes in ``src/docx/oxml/`` are thin
  subclasses of ``lxml.etree.ElementBase``. They are the XML; they don't wrap
  anything. They give element types a name (``CT_Footnote``, ``CT_Paragraph``,
  ``CT_R``), provide typed accessors for child elements and attributes, and
  enforce OOXML schema ordering when new children are inserted.

Looking at it from the other direction: every proxy object holds its
underlying XML element in a ``_element`` (often aliased ``_p``, ``_r``, ``_tc``,
etc.) attribute. Every such element is an instance of a ``CT_*`` class. The
part that owns it is reachable through its ancestors. You can always move
between layers — the split is organizational, not a security boundary.


A concrete pair: ``Footnote`` and ``CT_Footnote``
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Here is a proxy class and the oxml element it wraps, reduced to the essentials.
The oxml class describes the ``<w:footnote>`` element, and the proxy class
gives application code a friendly API over it::

    # src/docx/oxml/footnotes.py
    from docx.oxml.xmlchemy import (
        BaseOxmlElement, RequiredAttribute, ZeroOrMore, ZeroOrOne,
    )
    from docx.oxml.simpletypes import ST_DecimalNumber

    class CT_Footnote(BaseOxmlElement):
        """``<w:footnote>`` element."""
        pPr = ZeroOrOne("w:pPr", successors=("w:r",))
        r = ZeroOrMore("w:r", successors=())
        id = RequiredAttribute("w:id", ST_DecimalNumber)

    # src/docx/footnotes.py
    class Footnote(BlockItemContainer):
        """Proxy for a single ``<w:footnote>`` element."""

        @property
        def footnote_id(self):
            return self._element.id

Two things are worth noting:

* The proxy's ``footnote_id`` delegates straight to ``self._element.id``.
  The proxy adds no storage; it only translates attribute access into
  operations on the underlying element.

* The ``CT_Footnote`` class says nothing about "the Python API"; it is
  exclusively a description of XML shape. A different proxy could be layered
  on top of the same element class without any changes to oxml.


``BaseOxmlElement``, ``ZeroOrOne``, ``ZeroOrMore``, and successors
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The oxml layer is built with an internal helper package called
``xmlchemy``. It provides descriptors that turn child-element and attribute
declarations on a ``CT_*`` class into Pythonic accessors:

* ``ZeroOrOne(tag, successors=(...))`` — declares an optional single child.
  It generates a read-only attribute with that name, plus
  ``_add_<tag>()``, ``get_or_add_<tag>()``, ``_remove_<tag>()``, and
  ``_insert_<tag>()`` helpers.

* ``ZeroOrMore(tag, successors=(...))`` — declares a repeating child.
  It generates a ``<tag>_lst`` property returning a list, plus
  ``add_<tag>()`` and ``_insert_<tag>()`` helpers.

* ``OneAndOnlyOne`` / ``OneOrMore`` — variants with different cardinality
  semantics.

* ``RequiredAttribute`` / ``OptionalAttribute`` — typed attribute descriptors
  that validate values through an ``ST_*`` simpletype.

The ``successors`` tuple is important. OOXML is position-sensitive: a
``<w:pPr>`` that appears *after* the runs of a paragraph is not the same
thing as one that appears before them — Word will reject or silently mangle
malformed ordering. ``successors`` names the sibling tags that, if present,
must come *after* the element being inserted. ``xmlchemy`` uses it to pick a
correct insertion point when adding children.

Getting ``successors`` right therefore requires consulting the schema.
The canonical source for element ordering lives in the ``spec/`` folder at
the repository root:

* ``spec/xsd/wml.xsd`` — WordprocessingML (paragraphs, runs, tables, sections,
  footnotes, comments, etc.).
* ``spec/xsd/dml-wordprocessingDrawing.xsd`` — inline and anchor drawing
  wrappers used for images and shapes inside a Word document.
* ``spec/xsd/shared-math.xsd`` — Office Math (OMML).
* ``spec/rnc/*.rnc`` — RELAX NG Compact equivalents of the same grammars,
  substantially easier to read than the XSDs when you only want the ordering.

When the XSD and observed Word behaviour disagree, treat Word's behaviour
as authoritative for interoperability and the XSD as authoritative for what
the spec permits.


Registering elements and relationships
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Declaring a ``CT_*`` class is not enough on its own — lxml needs to know
that the tag ``w:footnote`` should be parsed into an instance of
``CT_Footnote``. That mapping is installed by ``register_element_cls`` in
``src/docx/oxml/__init__.py``::

    register_element_cls("w:footnote", CT_Footnote)

Without that line, parsing a footnote element would give you a generic lxml
element and the descriptors on ``CT_Footnote`` would never run.

Adjacent constants live in ``src/docx/opc/constants.py``. That module
defines both content types (``CT.WML_FOOTNOTES``, ``CT.WML_COMMENTS``, etc.)
and relationship types (``RT.FOOTNOTES``, ``RT.COMMENTS``) used when a part
attaches itself to the package. New part classes — defined under
``src/docx/parts/`` and registered on ``PartFactory.part_type_for`` in
``src/docx/__init__.py`` — reach for these constants rather than hard-coding
URIs.


Reaching into ``_element``: when (and when not) to do it
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Because |docx|'s proxies are deliberately thin, the escape hatch is simple:
every proxy exposes its underlying oxml element, and every oxml element is
itself an lxml element. If the high-level API does not yet cover a feature
you need, you can always drop down::

    # paragraph.pPr is a CT_PPr (lxml element). From here, any
    # WordprocessingML that can appear under w:pPr is reachable.
    pPr = paragraph._p.get_or_add_pPr()
    shd = pPr.makeelement(qn("w:shd"), {qn("w:fill"): "FFFF00"})
    pPr.append(shd)

This is fully supported — it is the same API the library's own proxies are
implemented with. A few rules of thumb:

* **Prefer the public API where it exists.** The proxies exist precisely
  because OOXML has many sharp edges (ordering, reserved IDs, namespace
  aliases, schema transitions). When they cover your case, use them.

* **Drop to oxml for feature gaps.** If you need a paragraph property
  python-docx does not yet surface — ruby text, conditional formatting,
  a Word 2013+ extension — call ``paragraph._p`` (or ``.element``) and
  manipulate the tree directly. This is a legitimate and common pattern.

* **Treat oxml as semi-public.** ``CT_*`` class names, attribute descriptors,
  and the ``_p`` / ``_r`` / ``_tc`` / ``_element`` accessors are stable
  enough to build on. Deep internals of ``xmlchemy`` (the descriptor
  implementation itself) are not.

* **Respect schema ordering.** When you insert a new child from oxml, use
  ``_insert_<tag>()`` on its parent (if one is generated) rather than a bare
  ``append()``. That insertion uses the ``successors`` tuple described above
  and keeps Word happy.

* **Namespace everything.** Use ``docx.oxml.ns.qn("w:tag")`` to build
  Clark-notation tag names — never hard-coded strings — so that the right
  namespace URI is always attached.

With that shape in mind, the rest of the user guide — individual topics like
comments, footnotes, numbering, or track changes — can be read as a tour of
the proxy APIs layered on top of this same three-tier foundation.
