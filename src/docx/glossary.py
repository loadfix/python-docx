"""|Glossary| proxy and related classes for the glossary-document part.

Provides read-only access to the building blocks (AutoText / Quick Parts /
cover pages) stored in ``word/glossary/document.xml``. Access via
:attr:`docx.document.Document.glossary`, which returns a :class:`Glossary`
when the document has a ``glossaryDocument`` relationship, or |None|
otherwise.

Building blocks are exposed as :class:`BuildingBlock` objects, each with a
:class:`BuildingBlockCategory` proxy for the OOXML
``w:docPart/w:docPartPr/w:category`` element. The block's content paragraphs
and tables are exposed via the standard block-item container API.
"""

from __future__ import annotations

import uuid as _uuid
from collections.abc import Iterator
from typing import TYPE_CHECKING, Union

from docx.blkcntnr import BlockItemContainer
from docx.enum.text import WD_BUILDING_BLOCK_GALLERY
from docx.shared import ElementProxy

if TYPE_CHECKING:
    import docx.types as t
    from docx.oxml.glossary import (
        CT_DocPart,
        CT_DocPartCategory,
        CT_GlossaryDocument,
    )
    from docx.parts.glossary import GlossaryPart
    from docx.table import Table
    from docx.text.paragraph import Paragraph


class Glossary(ElementProxy):
    """Proxy for the ``w:glossaryDocument`` root of the glossary part.

    Iterable: iterating yields a |BuildingBlock| for each ``w:docPart``
    child in document order. Supports ``len()`` and ``glossary[name]``
    lookup by building-block name. Read-only.

    .. versionadded:: 2026.05.0
    """

    def __init__(
        self,
        element: CT_GlossaryDocument,
        parent: GlossaryPart | t.ProvidesXmlPart | None = None,
    ):
        super().__init__(element, parent)
        self._glossary_elm = element
        self._glossary_part = parent

    def __iter__(self) -> Iterator[BuildingBlock]:
        """Yield a |BuildingBlock| for each ``w:docPart`` in this glossary."""
        return iter(self.building_blocks)

    def __len__(self) -> int:
        """Number of building blocks in this glossary."""
        return len(self._glossary_elm.docPart_lst)

    def __getitem__(self, name: str) -> BuildingBlock:
        """Return the |BuildingBlock| whose name is `name`.

        Raises |KeyError| when no building block with that name is present.
        Name comparison is exact (case-sensitive); the first match in
        document order wins when names collide.
        """
        for block in self.building_blocks:
            if block.name == name:
                return block
        raise KeyError(name)

    @property
    def building_blocks(self) -> list[BuildingBlock]:
        """List of |BuildingBlock| objects, one per ``w:docPart``, in order.

        .. versionadded:: 2026.05.0
        """
        return [
            BuildingBlock(doc_part, self._glossary_part)
            for doc_part in self._glossary_elm.docPart_lst
        ]

    def add_building_block(
        self,
        name: str,
        category: str = "General",
        gallery: WD_BUILDING_BLOCK_GALLERY | str = (
            WD_BUILDING_BLOCK_GALLERY.QUICK_PARTS
        ),
        content: Union["Paragraph", str, None] = None,
    ) -> BuildingBlock:
        """Add a new building block and return its :class:`BuildingBlock`.

        `name` becomes ``w:docPartPr/w:name/@w:val`` — the display name used
        by Word when listing Quick Parts. `category` and `gallery` populate
        ``w:docPartPr/w:category`` (``w:name`` and ``w:gallery``
        respectively); the default gallery is ``QUICK_PARTS`` since that is
        the most common bucket for user-authored snippets. `gallery` may be
        passed as a :class:`WD_BUILDING_BLOCK_GALLERY` enum member or as its
        raw XML string (e.g. ``"coverPg"``).

        `content` populates ``w:docPartBody``. A ``str`` is wrapped in a
        single paragraph with a single run; a :class:`docx.text.paragraph.Paragraph`
        has its underlying ``<w:p>`` element appended directly (making the
        building block "take ownership" — the caller's paragraph now lives
        in the glossary part). |None| produces an empty body.

        A fresh ``w:guid`` is generated for every new block — Word uses this
        to disambiguate building blocks that share a name across galleries.

        .. versionadded:: 2026.05.10
        """
        if isinstance(gallery, WD_BUILDING_BLOCK_GALLERY):
            gallery_xml = gallery.xml_value
        else:
            gallery_xml = gallery

        doc_parts = self._glossary_elm.get_or_add_docParts()
        doc_part = doc_parts.add_docPart()
        pr = doc_part.get_or_add_docPartPr()
        pr.set_name(name)
        cat = pr.get_or_add_category()
        cat.set_name(category)
        cat.set_gallery(gallery_xml)
        pr.set_guid("{%s}" % _uuid.uuid4())
        body = doc_part.get_or_add_docPartBody()

        if isinstance(content, str):
            p = body.add_p()
            # -- append a run with the supplied text --
            from docx.oxml.text.run import CT_R
            run_elm: CT_R = p.add_r()
            run_elm.text = content
        elif content is not None:
            # -- a Paragraph proxy — detach its element and append --
            body.append(content._p)  # type: ignore[attr-defined]

        return BuildingBlock(doc_part, self._glossary_part)

    def remove_building_block(self, name: str) -> bool:
        """Remove the first building block whose name is `name`.

        Returns ``True`` when a block was removed, ``False`` when no match
        exists. The first match in document order wins when names collide.
        Name comparison is exact and case-sensitive.

        .. versionadded:: 2026.05.10
        """
        docParts = self._glossary_elm.docParts
        if docParts is None:
            return False
        for doc_part in list(docParts.docPart_lst):
            pr = doc_part.docPartPr
            if pr is None:
                continue
            if pr.name_val == name:
                docParts.remove(doc_part)
                return True
        return False

    def by_category(
        self,
        gallery: WD_BUILDING_BLOCK_GALLERY | str | None = None,
        category_name: str | None = None,
    ) -> list[BuildingBlock]:
        """Return building blocks filtered by gallery and/or `category_name`.

        Either argument may be omitted; when both are provided, results
        intersect. `gallery` may be a :class:`WD_BUILDING_BLOCK_GALLERY`
        member or a raw XML string (e.g. ``"quickParts"``); raw strings are
        compared as-is to the underlying gallery value. When both arguments
        are |None|, every building block is returned (equivalent to
        :attr:`building_blocks`). Comparison for `category_name` is exact
        and case-sensitive.

        .. versionadded:: 2026.05.0
        """
        if isinstance(gallery, WD_BUILDING_BLOCK_GALLERY):
            gallery_xml: str | None = gallery.xml_value
        else:
            gallery_xml = gallery

        result: list[BuildingBlock] = []
        for block in self.building_blocks:
            cat = block.category
            if gallery_xml is not None and cat.gallery != gallery_xml:
                continue
            if category_name is not None and cat.category_name != category_name:
                continue
            result.append(block)
        return result

    @property
    def categories(self) -> list[BuildingBlockCategory]:
        """Unique |BuildingBlockCategory| objects across all building blocks.

        Deduplication is by the ``(gallery, category_name)`` pair — two
        categories with the same gallery value and same name count as one,
        regardless of which underlying ``w:category`` element they came
        from. Order preserves first-seen order in document traversal.
        Categories where both gallery and name are |None| are dropped.

        .. versionadded:: 2026.05.0
        """
        seen: set[tuple[str | None, str | None]] = set()
        result: list[BuildingBlockCategory] = []
        for block in self.building_blocks:
            cat = block.category
            if cat.gallery is None and cat.category_name is None:
                continue
            key = (cat.gallery, cat.category_name)
            if key in seen:
                continue
            seen.add(key)
            result.append(cat)
        return result

    @property
    def galleries(self) -> list[str]:
        """Unique gallery XML strings across all building blocks.

        Returns the raw ``w:val`` strings (e.g. ``"quickParts"``,
        ``"coverPg"``) in first-seen order. Use
        :meth:`WD_BUILDING_BLOCK_GALLERY.from_xml_safe` to decode individual
        values. Building blocks with no gallery are skipped.

        .. versionadded:: 2026.05.0
        """
        seen: set[str] = set()
        result: list[str] = []
        for block in self.building_blocks:
            gallery = block.category.gallery
            if gallery is None or gallery in seen:
                continue
            seen.add(gallery)
            result.append(gallery)
        return result


class BuildingBlock(BlockItemContainer):
    """Proxy for a single ``w:docPart`` (a building block) in the glossary.

    A building block has metadata (name, category, description, GUID) and a
    body composed of block items (paragraphs and tables). The body is
    accessed through the :class:`BlockItemContainer` API — the element
    passed to the base class is the ``w:docPartBody`` child, which carries
    the paragraphs and tables.

    When the building block has no ``w:docPartBody`` child the
    :attr:`paragraphs` and :attr:`tables` properties return empty lists.

    .. versionadded:: 2026.05.0
    """

    def __init__(
        self,
        doc_part: CT_DocPart,
        parent: GlossaryPart | t.ProvidesXmlPart | None = None,
    ):
        # -- ``BlockItemContainer`` expects the *story* element (the one
        # -- holding the paragraphs and tables); for a building block that
        # -- is the ``w:docPartBody`` child. When the body is absent we
        # -- pass the ``w:docPart`` itself — the empty-list overrides
        # -- below keep callers sane for that case.
        body = doc_part.docPartBody
        element_for_container = body if body is not None else doc_part
        super().__init__(element_for_container, parent)  # type: ignore[arg-type]
        self._doc_part = doc_part
        self._body = body

    @property
    def name(self) -> str | None:
        """The value of ``w:docPartPr/w:name/@w:val`` — the block's name.

        |None| when ``w:docPartPr`` is absent, when the ``w:name`` child is
        missing, or when its ``w:val`` attribute is not present.

        .. versionadded:: 2026.05.0
        """
        pr = self._doc_part.docPartPr
        if pr is None:
            return None
        return pr.name_val

    @property
    def category(self) -> BuildingBlockCategory:
        """A |BuildingBlockCategory| for this block's ``w:category``.

        Always returns a proxy — when the underlying ``w:category`` element
        is missing, the returned proxy exposes |None| for every slot.

        .. versionadded:: 2026.05.0
        """
        pr = self._doc_part.docPartPr
        category_elm = pr.category if pr is not None else None
        return BuildingBlockCategory(category_elm)

    @property
    def description(self) -> str | None:
        """The value of ``w:docPartPr/w:description/@w:val``, or |None|.

        .. versionadded:: 2026.05.0
        """
        pr = self._doc_part.docPartPr
        if pr is None:
            return None
        return pr.description_val

    @property
    def guid(self) -> str | None:
        """The value of ``w:docPartPr/w:guid/@w:val``, or |None|.

        .. versionadded:: 2026.05.0
        """
        pr = self._doc_part.docPartPr
        if pr is None:
            return None
        return pr.guid_val

    @property
    def uuid(self) -> str | None:
        """Alias of :attr:`guid`, returning the ``w:guid`` ``w:val`` slot.

        The OOXML element is named ``w:guid`` so that remains the canonical
        spelling; this property is provided for callers who prefer the
        vendor-neutral name.

        .. versionadded:: 2026.05.10
        """
        return self.guid

    @property
    def types(self) -> list[str]:
        """List of raw ``w:types/w:type/@w:val`` strings for this block.

        Empty when ``w:docPartPr`` or ``w:types`` is absent, or when the
        ``w:types`` element has no children.

        .. versionadded:: 2026.05.10
        """
        pr = self._doc_part.docPartPr
        if pr is None or pr.types is None:
            return []
        return pr.types.values

    @property
    def type(self) -> str | None:
        """The first ``w:types/w:type/@w:val`` string, or |None| when absent.

        Convenience for building blocks that declare a single type (the
        common case — Word writes one ``w:type`` child with a value like
        ``"autoTxt"``). Use :attr:`types` to see every declared type.

        .. versionadded:: 2026.05.10
        """
        type_values = self.types
        return type_values[0] if type_values else None

    @property
    def behaviors(self) -> list[str]:
        """List of raw ``w:behaviors/w:behavior/@w:val`` strings.

        Empty when ``w:docPartPr`` or ``w:behaviors`` is absent.

        .. versionadded:: 2026.05.10
        """
        pr = self._doc_part.docPartPr
        if pr is None or pr.behaviors is None:
            return []
        return pr.behaviors.values

    @property
    def content_paragraphs(self) -> list[Paragraph]:
        """Alias of :attr:`paragraphs`.

        The task-spec vocabulary surfaces "content paragraphs" as a
        distinct slot from the `BlockItemContainer` paragraphs property —
        they refer to the same ``w:docPartBody`` contents, but the alias
        keeps API callers aligned with the glossary-schema terminology.

        .. versionadded:: 2026.05.10
        """
        return self.paragraphs

    @property
    def paragraphs(self) -> list[Paragraph]:
        """List of |Paragraph| objects in the building block's body.

        Returns an empty list when the block has no ``w:docPartBody`` child.

        .. versionadded:: 2026.05.0
        """
        if self._body is None:
            return []
        return super().paragraphs

    @property
    def tables(self) -> list[Table]:
        """List of |Table| objects in the building block's body.

        Returns an empty list when the block has no ``w:docPartBody`` child.

        .. versionadded:: 2026.05.0
        """
        if self._body is None:
            return []
        return super().tables


class BuildingBlockCategory:
    """Read-only view over a building block's ``w:category`` element.

    Exposes the category name (``w:category/w:name/@w:val``) and gallery
    (``w:category/w:gallery/@w:val``). Both return |None| when the
    underlying element or attribute is missing. Equality is by
    ``(gallery, category_name)`` so categories with identical slots are
    interchangeable — convenient for set-based deduplication.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, category_elm: CT_DocPartCategory | None):
        self._category_elm = category_elm

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, BuildingBlockCategory):
            return NotImplemented
        return (
            self.gallery == other.gallery
            and self.category_name == other.category_name
        )

    def __hash__(self) -> int:
        return hash((self.gallery, self.category_name))

    def __repr__(self) -> str:
        return (
            f"BuildingBlockCategory(gallery={self.gallery!r}, "
            f"category_name={self.category_name!r})"
        )

    @property
    def category_name(self) -> str | None:
        """The value of ``w:category/w:name/@w:val``, or |None| when absent.

        .. versionadded:: 2026.05.0
        """
        if self._category_elm is None:
            return None
        return self._category_elm.name_val

    @property
    def gallery(self) -> str | None:
        """The value of ``w:category/w:gallery/@w:val``, or |None| when absent.

        .. versionadded:: 2026.05.0
        """
        if self._category_elm is None or self._category_elm.gallery is None:
            return None
        return self._category_elm.gallery.val

    @property
    def gallery_value(self) -> WD_BUILDING_BLOCK_GALLERY | None:
        """The gallery as a |WD_BUILDING_BLOCK_GALLERY| member, or |None|.

        |None| when the gallery slot is missing, or when its value is not
        one of the well-known Word galleries modelled by the enum. Use
        :attr:`gallery` to get the raw string for unknown values.

        .. versionadded:: 2026.05.0
        """
        return WD_BUILDING_BLOCK_GALLERY.from_xml_safe(self.gallery)
