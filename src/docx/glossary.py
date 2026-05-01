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
        """List of |BuildingBlock| objects, one per ``w:docPart``, in order."""
        return [
            BuildingBlock(doc_part, self._glossary_part)
            for doc_part in self._glossary_elm.docPart_lst
        ]


class BuildingBlock(BlockItemContainer):
    """Proxy for a single ``w:docPart`` (a building block) in the glossary.

    A building block has metadata (name, category, description, GUID) and a
    body composed of block items (paragraphs and tables). The body is
    accessed through the :class:`BlockItemContainer` API — the element
    passed to the base class is the ``w:docPartBody`` child, which carries
    the paragraphs and tables.

    When the building block has no ``w:docPartBody`` child the
    :attr:`paragraphs` and :attr:`tables` properties return empty lists.
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
        """
        pr = self._doc_part.docPartPr
        category_elm = pr.category if pr is not None else None
        return BuildingBlockCategory(category_elm)

    @property
    def description(self) -> str | None:
        """The value of ``w:docPartPr/w:description/@w:val``, or |None|."""
        pr = self._doc_part.docPartPr
        if pr is None:
            return None
        return pr.description_val

    @property
    def guid(self) -> str | None:
        """The value of ``w:docPartPr/w:guid/@w:val``, or |None|."""
        pr = self._doc_part.docPartPr
        if pr is None:
            return None
        return pr.guid_val

    @property
    def paragraphs(self) -> list[Paragraph]:
        """List of |Paragraph| objects in the building block's body.

        Returns an empty list when the block has no ``w:docPartBody`` child.
        """
        if self._body is None:
            return []
        return super().paragraphs

    @property
    def tables(self) -> list[Table]:
        """List of |Table| objects in the building block's body.

        Returns an empty list when the block has no ``w:docPartBody`` child.
        """
        if self._body is None:
            return []
        return super().tables


class BuildingBlockCategory:
    """Read-only view over a building block's ``w:category`` element.

    Exposes the category name (``w:category/w:name/@w:val``) and gallery
    (``w:category/w:gallery/@w:val``). Both return |None| when the
    underlying element or attribute is missing.
    """

    def __init__(self, category_elm: CT_DocPartCategory | None):
        self._category_elm = category_elm

    @property
    def category_name(self) -> str | None:
        """The value of ``w:category/w:name/@w:val``, or |None| when absent."""
        if self._category_elm is None:
            return None
        return self._category_elm.name_val

    @property
    def gallery(self) -> str | None:
        """The value of ``w:category/w:gallery/@w:val``, or |None| when absent."""
        if self._category_elm is None or self._category_elm.gallery is None:
            return None
        return self._category_elm.gallery.val
