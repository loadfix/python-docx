"""Custom element classes for inline CustomXml container elements.

``CustomXml*`` containers wrap document content with a user-defined XML
element and a set of attribute overrides. There are four flavors of
container, mirroring the WordprocessingML content-group hierarchy:

- ``w:customXml`` as block-level wrapper (``CT_CustomXmlBlock``)
- ``w:customXml`` inside a table row (``CT_CustomXmlRow``)
- ``w:customXml`` inside a table-row cell (``CT_CustomXmlCell``)
- ``w:customXml`` inline (``CT_CustomXmlRun``)

All four share the same outer shape — a required ``@w:element`` attribute
(the user-defined element local-name), an optional ``@w:uri`` attribute (the
namespace URI), an optional ``w:customXmlPr`` child, and zero-or-more
content children of the flavor-appropriate group.

The property element ``w:customXmlPr`` (``CT_CustomXmlPr``) contains an
optional placeholder plus zero or more ``w:attr`` custom-attribute records.

Note: ``w:customXml`` reuses the same local name across all four flavors;
lxml registers a single element class per QName, so :func:`register_element_cls`
in :mod:`docx.oxml` binds it to :class:`CT_CustomXmlBlock` by default. The
block class's grammar is the most permissive (block-level content) and parses
run/row/cell content positionally, so inline and table-scoped trees round-trip
faithfully through the block flavor. The remaining three CT classes are
declared for explicit construction in tests and for downstream consumers that
instantiate a flavor-specific container programmatically.

.. versionadded:: 2026.05.10
"""

from __future__ import annotations

from docx.oxml.simpletypes import ST_String, XsdString
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


# ---------------------------------------------------------------------------
# CT_Attr — custom attribute pair (``w:attr``)
# ---------------------------------------------------------------------------


class CT_Attr(BaseOxmlElement):
    """`w:attr` — one custom-attribute entry inside ``w:customXmlPr``.

    Carries an optional ``@w:uri`` (namespace URI of the custom attribute),
    a required ``@w:name`` (local name), and a required ``@w:val`` (value).
    """

    uri: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:uri", ST_String
    )
    name: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:name", ST_String
    )
    val: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


# ---------------------------------------------------------------------------
# CT_CustomXmlPr — properties element (``w:customXmlPr``)
# ---------------------------------------------------------------------------


class CT_CustomXmlPr(BaseOxmlElement):
    """`w:customXmlPr` — properties of the enclosing ``w:customXml`` container.

    Holds an optional ``w:placeholder`` (a reference to a placeholder text
    style by name) and zero or more ``w:attr`` custom-attribute records.
    """

    _tag_seq = ("w:placeholder", "w:attr")

    placeholder: "BaseOxmlElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:placeholder", successors=("w:attr",)
    )
    attr = ZeroOrMore("w:attr", successors=())

    del _tag_seq


# ---------------------------------------------------------------------------
# CustomXml containers — one per content-group flavor
# ---------------------------------------------------------------------------


class CT_CustomXmlBlock(BaseOxmlElement):
    """`w:customXml` (block-level) — wraps block-level content with a
    user-defined XML element.

    The ``@w:element`` attribute is the user-defined element local name; the
    optional ``@w:uri`` attribute is its namespace URI. An optional
    ``w:customXmlPr`` child carries the placeholder / attribute metadata.
    """

    customXmlPr: "CT_CustomXmlPr | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:customXmlPr", successors=()
    )

    uri: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:uri", ST_String
    )
    element: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:element", XsdString
    )


class CT_CustomXmlRun(BaseOxmlElement):
    """`w:customXml` (inline) — wraps inline run-level content with a
    user-defined XML element.

    Identical outer shape to :class:`CT_CustomXmlBlock` but with the ``EG_PContent``
    group as allowed children (runs, fields, hyperlinks, sub-docs).
    """

    customXmlPr: "CT_CustomXmlPr | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:customXmlPr", successors=()
    )

    uri: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:uri", ST_String
    )
    element: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:element", XsdString
    )


class CT_CustomXmlRow(BaseOxmlElement):
    """`w:customXml` (table-row) — wraps one or more ``w:tr`` elements with
    a user-defined XML element.

    Identical outer shape to :class:`CT_CustomXmlBlock` but intended to
    live directly inside ``w:tbl`` and to contain the ``EG_ContentRowContent``
    group (``w:tr`` plus SDT/tracked-change markers that carry rows).
    """

    customXmlPr: "CT_CustomXmlPr | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:customXmlPr", successors=()
    )

    uri: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:uri", ST_String
    )
    element: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:element", XsdString
    )


class CT_CustomXmlCell(BaseOxmlElement):
    """`w:customXml` (table-cell) — wraps one or more ``w:tc`` elements with
    a user-defined XML element.

    Identical outer shape to :class:`CT_CustomXmlBlock` but intended to live
    directly inside ``w:tr`` and to contain the ``EG_ContentCellContent``
    group (``w:tc`` plus SDT/tracked-change markers that carry cells).
    """

    customXmlPr: "CT_CustomXmlPr | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:customXmlPr", successors=()
    )

    uri: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:uri", ST_String
    )
    element: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:element", XsdString
    )
