"""Custom element classes related to the ``word/glossary/document.xml`` part.

The glossary document part carries the "building blocks" (AutoText / Quick
Parts / cover pages, etc.) that ship with a document. Each building block is
a ``w:docPart`` inside ``w:docParts``, containing a ``w:docPartPr`` metadata
block and a ``w:docPartBody`` holding the block's content paragraphs and
tables.

Only the pieces exposed through :class:`docx.glossary.Glossary` are modelled
at this pass — the building-block content is treated as a generic story
(paragraphs plus tables). Creation of new building blocks is intentionally
out of scope; the proxy layer is read-only.

``w:name`` is already registered globally as :class:`CT_String` (it's used as
a simple ``@w:val`` carrier under a handful of WML parents — style names,
footer references, etc.), so the glossary-specific classes below read the
``w:val`` attribute of their ``w:name`` / ``w:description`` / ``w:guid``
children directly via ``xpath`` rather than via a typed child accessor.
That keeps the global ``w:name`` registration intact.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.oxml.ns import nsmap, qn
from docx.oxml.simpletypes import ST_OnOff, ST_String
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P


# Schema order for ``w:docPartPr`` children per ECMA-376-1 CT_DocPartPr.
# Used when inserting a newly-created child to preserve a valid sequence.
_DOC_PART_PR_CHILD_ORDER: tuple[str, ...] = (
    "w:name",
    "w:style",
    "w:category",
    "w:types",
    "w:behaviors",
    "w:description",
    "w:guid",
)


def _set_w_val_child(
    parent: BaseOxmlElement,
    child_tag: str,
    val: str,
    predecessors: tuple[str, ...],  # unused — kept for call-site readability
) -> BaseOxmlElement:
    """Set `child_tag`'s ``w:val`` under `parent`, creating the child if absent.

    When `child_tag` is missing, a new element is created via
    ``makeelement`` and inserted in schema order according to
    :data:`_DOC_PART_PR_CHILD_ORDER`. Returns the target child element.
    """
    del predecessors  # noqa: F841 — accepted for readability at call sites
    found = parent.xpath(f"./{child_tag}[1]")
    if found:
        child = found[0]
    else:
        # -- build and insert in schema order --
        nsmap_for_create = {
            prefix: uri for prefix, uri in nsmap.items() if prefix == "w"
        }
        child = parent.makeelement(qn(child_tag), nsmap=nsmap_for_create)
        # Determine insertion point: first existing child whose schema index
        # is greater than this one's.
        try:
            our_idx = _DOC_PART_PR_CHILD_ORDER.index(child_tag)
        except ValueError:
            our_idx = len(_DOC_PART_PR_CHILD_ORDER)
        insert_at = len(parent)
        for i, existing in enumerate(parent):
            try:
                # ``existing.tag`` is a Clark-notation string; round-trip via
                # the local-name map to recover a ``w:x`` short tag for lookup.
                local = existing.tag.rsplit("}", 1)[-1]
                existing_short = f"w:{local}"
                existing_idx = _DOC_PART_PR_CHILD_ORDER.index(existing_short)
            except ValueError:
                continue
            if existing_idx > our_idx:
                insert_at = i
                break
        parent.insert(insert_at, child)
    child.set(qn("w:val"), val)
    return child


def _val_of(element: BaseOxmlElement | None) -> str | None:
    """Return the ``w:val`` attribute of `element`, or |None|.

    Returns |None| when `element` is |None| or the attribute is not present.
    Avoids taking a typed ``CT_String`` route because several of the glossary
    children share a local-name (``w:name``) with other WML elements whose
    registered class is ``CT_String`` — this helper keeps the reader tolerant
    of both.
    """
    if element is None:
        return None
    return element.get(qn("w:val"))


class CT_DocPartName(BaseOxmlElement):
    """``<w:name>`` child of ``w:docPartPr`` — the building block's name.

    Carries the block's display name in its ``w:val`` attribute. The global
    ``w:name`` tag is registered as :class:`docx.oxml.shared.CT_String`
    (because other WML schemas use the same local name); this class exists as
    a typed shape for constructing the element with a known attribute name
    and for documentation purposes. Read access still flows through the
    ``xpath`` helper on :class:`CT_DocPartPr` to remain compatible with the
    globally-registered ``CT_String``.
    """

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_DocPartBehavior(BaseOxmlElement):
    """``<w:behavior>`` child of ``w:behaviors`` — a single behavior flag.

    The ``w:val`` attribute names the behavior (``"content"``, ``"p"``,
    ``"pg"``). Not registered as the global ``w:behavior`` class — the
    element is read via xpath inside :class:`CT_DocPartBehaviors` to avoid
    colliding with any future global registrations. Instances produced by
    ``makeelement`` on a parent still carry the ``val`` attribute shape.
    """

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_DocPartBehaviors(BaseOxmlElement):
    """``<w:behaviors>`` child of ``w:docPartPr`` — list of behavior flags.

    Contains zero-or-more ``w:behavior`` children. Word writes this element
    when the building block has behaviors other than the implicit "content".
    ``w:behavior`` is read via xpath to keep the global ``w:behavior`` tag
    open for other uses.
    """

    @property
    def values(self) -> list[str]:
        """List of ``w:val`` attributes of the child ``w:behavior`` elements.

        Missing ``w:val`` attributes are skipped; empty list when there are
        no ``w:behavior`` children.
        """
        return [
            val
            for child in self.xpath("./w:behavior")
            if (val := child.get(qn("w:val"))) is not None
        ]

    def add_behavior(self, val: str) -> None:
        """Append a new ``w:behavior`` child with ``@w:val=<val>``."""
        child = self.makeelement(qn("w:behavior"), nsmap={"w": nsmap["w"]})
        child.set(qn("w:val"), val)
        self.append(child)


class CT_DocPartType(BaseOxmlElement):
    """``<w:type>`` child of ``w:types`` — a single type flag.

    The ``w:val`` attribute names the building-block type (``"autoTxt"``,
    ``"toolbar"``, etc.). See :class:`CT_DocPartTypes` for why this class
    is not bound as the global ``w:type`` registration (``CT_SectType``
    owns that).
    """

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_DocPartTypes(BaseOxmlElement):
    """``<w:types>`` child of ``w:docPartPr`` — list of type flags.

    Contains zero-or-more ``w:type`` children plus an ``@w:all`` boolean
    attribute that indicates the block is usable in every context.
    ``w:type`` children are read via xpath because the global ``w:type``
    tag is registered as ``CT_SectType`` for sections.
    """

    all: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:all", ST_OnOff
    )

    @property
    def values(self) -> list[str]:
        """List of ``w:val`` attributes of the child ``w:type`` elements."""
        return [
            val
            for child in self.xpath("./w:type")
            if (val := child.get(qn("w:val"))) is not None
        ]

    def add_type(self, val: str) -> None:
        """Append a new ``w:type`` child with ``@w:val=<val>``."""
        child = self.makeelement(qn("w:type"), nsmap={"w": nsmap["w"]})
        child.set(qn("w:val"), val)
        self.append(child)


class CT_DocPartGallery(BaseOxmlElement):
    """``<w:gallery>`` child of ``w:category`` — the gallery slot for this block.

    The ``w:val`` attribute names the gallery (e.g. ``"quickParts"``,
    ``"coverPg"``). Modelled as a free-form string because the
    ``ST_DocPartGallery`` enumeration is broad and we only surface the value
    read-only.
    """

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_DocPartCategory(BaseOxmlElement):
    """``<w:category>`` child of ``w:docPartPr`` — the block's classification.

    Contains a name (``w:name``) and a gallery (``w:gallery``) — both are
    optional at the XML level but Word always writes them for first-class
    building blocks. ``w:name`` is read via ``xpath`` to avoid colliding
    with the global ``w:name`` registration elsewhere in the schema.
    """

    gallery: CT_DocPartGallery | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:gallery", successors=()
    )

    @property
    def name_val(self) -> str | None:
        """The value of ``w:category/w:name/@w:val``, or |None| when absent."""
        found = self.xpath("./w:name[1]")
        return _val_of(found[0] if found else None)

    def set_name(self, name: str) -> None:
        """Set ``w:category/w:name/@w:val`` to `name`, creating the child if absent.

        The new ``w:name`` is inserted before ``w:gallery`` to preserve the
        ECMA schema ordering.
        """
        found = self.xpath("./w:name[1]")
        if found:
            child = found[0]
        else:
            child = self.makeelement(qn("w:name"), nsmap={"w": nsmap["w"]})
            # -- insert before w:gallery if present, else at end --
            gallery_found = self.xpath("./w:gallery[1]")
            if gallery_found:
                gallery_found[0].addprevious(child)
            else:
                self.append(child)
        child.set(qn("w:val"), name)

    def set_gallery(self, gallery: str) -> None:
        """Set ``w:category/w:gallery/@w:val`` to `gallery`, creating it if absent."""
        g = self.gallery
        if g is None:
            g = self.get_or_add_gallery()
        g.val = gallery


class CT_DocPartPr(BaseOxmlElement):
    """``<w:docPartPr>`` element — metadata for a building block.

    Holds the block name (``w:name``), its category (``w:category``), a GUID
    (``w:guid``), a description (``w:description``), behaviors, types, and a
    ``w:style`` reference.
    """

    category: CT_DocPartCategory | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:category",
        successors=("w:types", "w:behaviors", "w:description", "w:guid"),
    )
    types: CT_DocPartTypes | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:types", successors=("w:behaviors", "w:description", "w:guid")
    )
    behaviors: CT_DocPartBehaviors | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:behaviors", successors=("w:description", "w:guid")
    )

    @property
    def name_val(self) -> str | None:
        """Value of ``w:docPartPr/w:name/@w:val``, or |None| when absent."""
        found = self.xpath("./w:name[1]")
        return _val_of(found[0] if found else None)

    @property
    def description_val(self) -> str | None:
        """Value of ``w:docPartPr/w:description/@w:val``, or |None| when absent."""
        found = self.xpath("./w:description[1]")
        return _val_of(found[0] if found else None)

    @property
    def guid_val(self) -> str | None:
        """Value of ``w:docPartPr/w:guid/@w:val``, or |None| when absent."""
        found = self.xpath("./w:guid[1]")
        return _val_of(found[0] if found else None)

    @property
    def style_val(self) -> str | None:
        """Value of ``w:docPartPr/w:style/@w:val``, or |None| when absent."""
        found = self.xpath("./w:style[1]")
        return _val_of(found[0] if found else None)

    # -- write helpers --------------------------------------------------

    def set_name(self, name: str) -> None:
        """Set the ``w:name`` child's ``w:val`` to `name`, creating it if absent."""
        _set_w_val_child(self, "w:name", name, predecessors=())

    def set_description(self, description: str) -> None:
        """Set the ``w:description`` child's ``w:val``, creating it if absent."""
        _set_w_val_child(
            self, "w:description", description, predecessors=("w:guid",)
        )

    def set_guid(self, guid: str) -> None:
        """Set the ``w:guid`` child's ``w:val``, creating it if absent."""
        _set_w_val_child(self, "w:guid", guid, predecessors=())

    # -- advanced-metadata helpers (R9-21) -----------------------------

    @property
    def docPartType_val(self) -> str | None:
        """The first ``w:types/w:type/@w:val``, or |None| when absent.

        ``w:docPartType`` is surfaced by :class:`docx.glossary.BuildingBlock`
        as a single-value slot; at the XML level WML carries a ``w:types``
        list with zero or more ``w:type`` children. This helper returns the
        first child's ``w:val``.
        """
        if self.types is None:
            return None
        values = self.types.values
        return values[0] if values else None

    def set_docPartType(self, val: str) -> None:
        """Set the first ``w:types/w:type/@w:val`` to `val`.

        Creates the ``w:types`` element and a single ``w:type`` child when
        absent; replaces the ``w:val`` on the existing first child otherwise.
        Other ``w:type`` siblings are preserved.
        """
        types = self.get_or_add_types()
        existing = types.xpath("./w:type")
        if existing:
            existing[0].set(qn("w:val"), val)
        else:
            types.add_type(val)

    def clear_docPartType(self) -> None:
        """Remove the ``w:types`` element if present."""
        types = self.types
        if types is not None:
            self.remove(types)

    def set_behaviors(self, values: list[str]) -> None:
        """Replace the ``w:behaviors`` children with one per entry in `values`.

        When `values` is empty the ``w:behaviors`` element is removed
        entirely so the XML round-trips cleanly.
        """
        beh = self.behaviors
        if beh is not None:
            self.remove(beh)
        if not values:
            return
        beh = self.get_or_add_behaviors()
        for v in values:
            beh.add_behavior(v)


class CT_DocPartBody(BaseOxmlElement):
    """``<w:docPartBody>`` element — container for a building block's content.

    A story-like container: holds paragraphs (``w:p``) and tables (``w:tbl``)
    in document order.
    """

    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    p_lst: list[CT_P]
    tbl_lst: list[CT_Tbl]

    @property
    def inner_content_elements(self) -> list[CT_P | CT_Tbl]:
        """All ``w:p`` and ``w:tbl`` elements in this doc-part body, in order."""
        return self.xpath("./w:p | ./w:tbl")


class CT_DocPart(BaseOxmlElement):
    """``<w:docPart>`` element — a single building block.

    Contains the metadata block (``w:docPartPr``) and the body
    (``w:docPartBody``) holding the block's content.
    """

    _tag_seq = ("w:docPartPr", "w:docPartBody")
    docPartPr: CT_DocPartPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:docPartPr", successors=("w:docPartBody",)
    )
    docPartBody: CT_DocPartBody | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:docPartBody", successors=()
    )
    del _tag_seq


class CT_DocParts(BaseOxmlElement):
    """``<w:docParts>`` element — container for the building blocks.

    Direct child of ``w:glossaryDocument``. Holds a flat list of
    ``w:docPart`` children.
    """

    docPart = ZeroOrMore("w:docPart", successors=())

    docPart_lst: list[CT_DocPart]


class CT_GlossaryDocument(BaseOxmlElement):
    """``<w:glossaryDocument>`` element — root of the glossary document part.

    Holds a single ``w:docParts`` container. Other children defined in the
    schema are not surfaced by this proxy.
    """

    docParts: CT_DocParts | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:docParts", successors=()
    )

    @property
    def docPart_lst(self) -> list[CT_DocPart]:
        """All ``w:docPart`` elements under this glossary document, in order.

        Returns an empty list when no ``w:docParts`` container is present or
        when it is empty.
        """
        docParts = self.docParts
        if docParts is None:
            return []
        return docParts.docPart_lst
