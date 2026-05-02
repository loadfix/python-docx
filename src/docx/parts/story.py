"""|StoryPart| and related objects."""

from __future__ import annotations

import io
import os
from typing import IO, TYPE_CHECKING, Iterator, cast

from docx.image.constants import MIME_TYPE
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import XmlPart
from docx.oxml.shape import CT_Anchor, CT_Inline
from docx.shared import Length, lazyproperty

if TYPE_CHECKING:
    from docx.enum.style import WD_STYLE_TYPE
    from docx.image.image import Image
    from docx.opc.part import Part
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.parts.document import DocumentPart
    from docx.styles.style import BaseStyle


class StoryPart(XmlPart):
    """Base class for story parts.

    A story part is one that can contain textual content, such as the document-part and
    header or footer parts. These all share content behaviors like `.paragraphs`,
    `.add_paragraph()`, `.add_table()` etc.
    """

    def get_or_add_image(
        self, image_descriptor: "str | os.PathLike[str] | IO[bytes]"
    ) -> tuple[str, Image]:
        """Return (rId, image) pair for image identified by `image_descriptor`.

        `rId` is the str key (often like "rId7") for the relationship between this story
        part and the image part, reused if already present, newly created if not.
        `image` is an |Image| instance providing access to the properties of the image,
        such as dimensions and image type.

        `image_descriptor` may be a ``str`` path, an :class:`os.PathLike` path, or a
        binary file-like object.

        .. versionchanged:: 1.3.0.dev0
           Accepts :class:`os.PathLike` path arguments.
        """
        if isinstance(image_descriptor, os.PathLike):
            image_descriptor = os.fspath(image_descriptor)
        package = self._package
        assert package is not None
        image_part = package.get_or_add_image_part(image_descriptor)
        rId = self.relate_to(image_part, RT.IMAGE)
        return rId, image_part.image

    def get_style(self, style_id: str | None, style_type: WD_STYLE_TYPE) -> BaseStyle:
        """Return the style in this document matching `style_id`.

        Returns the default style for `style_type` if `style_id` is |None| or does not
        match a defined style of `style_type`.
        """
        return self._document_part.get_style(style_id, style_type)

    def get_style_id(
        self, style_or_name: BaseStyle | str | None, style_type: WD_STYLE_TYPE
    ) -> str | None:
        """Return str style_id for `style_or_name` of `style_type`.

        Returns |None| if the style resolves to the default style for `style_type` or if
        `style_or_name` is itself |None|. Raises if `style_or_name` is a style of the
        wrong type or names a style not present in the document.
        """
        return self._document_part.get_style_id(style_or_name, style_type)

    def new_pic_inline(
        self,
        image_descriptor: "str | os.PathLike[str] | IO[bytes]",
        width: int | Length | None = None,
        height: int | Length | None = None,
    ) -> CT_Inline:
        """Return a newly-created `w:inline` element.

        The element contains the image specified by `image_descriptor` and is scaled
        based on the values of `width` and `height`.
        """
        rId, image = self.get_or_add_image(image_descriptor)
        cx, cy = image.scaled_dimensions(width, height)
        shape_id, filename = self.next_id, image.filename
        orientation = getattr(image, "orientation", None)

        if image.content_type == MIME_TYPE.SVG:
            return self._new_svg_pic_inline(
                shape_id, rId, filename, cx, cy, orientation=orientation
            )

        return CT_Inline.new_pic_inline(
            shape_id, rId, filename, cx, cy, orientation=orientation
        )

    def new_pic_anchor(
        self,
        image_descriptor: "str | os.PathLike[str] | IO[bytes]",
        width: int | Length | None = None,
        height: int | Length | None = None,
    ) -> CT_Anchor:
        """Return a newly-created `wp:anchor` element.

        The element contains the image specified by `image_descriptor` and is scaled
        based on the values of `width` and `height`.

        SVG images with a fallback PNG are not supported for floating images in this
        minimal implementation; SVG inputs are embedded via the PNG fallback path only
        through the standard picture relationship as for a regular raster image.
        """
        rId, image = self.get_or_add_image(image_descriptor)
        cx, cy = image.scaled_dimensions(width, height)
        shape_id, filename = self.next_id, image.filename
        orientation = getattr(image, "orientation", None)
        return CT_Anchor.new_pic_anchor(
            shape_id, rId, filename, cx, cy, orientation=orientation
        )

    def _new_svg_pic_inline(
        self,
        shape_id: int,
        svg_rId: str,
        filename: str,
        cx: Length,
        cy: Length,
        orientation: int | None = None,
    ) -> CT_Inline:
        """Return a `wp:inline` element for an SVG image with a PNG fallback."""
        fallback_png = self._generate_svg_fallback()
        fallback_stream = io.BytesIO(fallback_png)
        fallback_rId, _ = self.get_or_add_image(fallback_stream)
        return CT_Inline.new_svg_pic_inline(
            shape_id, fallback_rId, svg_rId, filename, cx, cy, orientation=orientation
        )

    @staticmethod
    def _generate_svg_fallback() -> bytes:
        """Return PNG bytes to use as SVG fallback.

        Generates a minimal 1x1 transparent PNG placeholder.
        """
        from docx.image.svg import generate_fallback_png

        return generate_fallback_png()

    @property
    def next_id(self) -> int:
        """Next available positive integer id value across all stories in the document.

        The value is determined by incrementing the maximum existing id value
        found in the main document body, every header part, and every footer
        part. Gaps in the existing id sequence are not filled. Spanning all
        story parts (rather than the current story only) avoids
        ``wp:docPr/@id`` collisions when images are added in a header/footer
        while the body — or another header/footer — already uses the same
        numeric id; Word interprets such collisions as the same drawing
        object.

        .. versionadded:: 1.3.0.dev0
        """
        used_ids: set[int] = set()
        for element in self._iter_story_elements():
            for id_str in element.xpath("//@id"):
                if id_str.isdigit():
                    used_ids.add(int(id_str))
        if not used_ids:
            return 1
        return max(used_ids) + 1

    def _iter_story_elements(self) -> Iterator[BaseOxmlElement]:
        """Yield the root XML element of each story part in the document.

        The current story's element is always yielded first; then the main
        document element (when this isn't already it) followed by each header
        and footer part's element, deduplicated. Errors resolving the document
        part or its related parts are swallowed so that unit tests can use a
        bare :class:`StoryPart` instance without a real package.
        """
        seen: set[int] = set()
        own = self._element
        if own is not None:
            seen.add(id(own))
            yield own

        doc_part = self._safe_document_part()
        if doc_part is None:
            return

        doc_element = cast(
            "BaseOxmlElement | None", getattr(doc_part, "_element", None)
        )
        if doc_element is not None and id(doc_element) not in seen:
            seen.add(id(doc_element))
            yield doc_element

        for reltype in (RT.HEADER, RT.FOOTER):
            for related in self._iter_related_parts(doc_part, reltype):
                related_element = cast(
                    "BaseOxmlElement | None", getattr(related, "_element", None)
                )
                if related_element is not None and id(related_element) not in seen:
                    seen.add(id(related_element))
                    yield related_element

    def _safe_document_part(self) -> Part | None:
        """Return the main |DocumentPart| or |None| if it can't be resolved.

        |StoryPart| instances constructed in tests may not have a package
        attached; guard those cases so ``next_id`` still produces a usable
        value from the local story alone.
        """
        try:
            return self._document_part
        except (AttributeError, AssertionError, KeyError):
            return None

    @staticmethod
    def _iter_related_parts(part: Part, reltype: str) -> Iterator[Part]:
        """Yield related parts of `part` matching `reltype`, tolerating a missing rels map."""
        rels = getattr(part, "rels", None)
        if rels is None:
            return
        for rel in rels.values():
            if getattr(rel, "is_external", False):
                continue
            if rel.reltype != reltype:
                continue
            try:
                yield rel.target_part
            except (KeyError, ValueError):  # pragma: no cover - defensive
                continue

    @lazyproperty
    def _document_part(self) -> DocumentPart:
        """|DocumentPart| object for this package."""
        package = self.package
        assert package is not None
        return cast("DocumentPart", package.main_document_part)
