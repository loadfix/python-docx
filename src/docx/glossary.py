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

from typing import TYPE_CHECKING
from collections.abc import Iterator

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
