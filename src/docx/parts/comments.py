"""Contains comments added to the document."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from docx.comments import Comments
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.oxml.comments import CT_Comments
from docx.oxml.parser import parse_xml
from docx.package import Package
from docx.parts.comments_extended import CommentsExtendedPart
from docx.parts.story import StoryPart

if TYPE_CHECKING:
    from docx.oxml.comments import CT_Comments
    from docx.package import Package


class CommentsPart(StoryPart):
    """Container part for comments added to the document."""

    def __init__(
        self, partname: PackURI, content_type: str, element: CT_Comments, package: Package
    ):
        super().__init__(partname, content_type, element, package)
        self._comments = element

    @property
    def comments(self) -> Comments:
        """A |Comments| proxy object for the `w:comments` root element of this part."""
        return Comments(self._comments, self)

    @classmethod
    def default(cls, package: Package) -> Self:
        """A newly created comments part, containing a default empty `w:comments` element."""
        partname = PackURI("/word/comments.xml")
        content_type = CT.WML_COMMENTS
        element = cast("CT_Comments", parse_xml(cls._default_comments_xml()))
        return cls(partname, content_type, element, package)

    @classmethod
    def _default_comments_xml(cls) -> bytes:
        """A byte-string containing XML for a default comments part."""
        return (Path(__file__).parent.parent / "templates" / "default-comments.xml").read_bytes()

    # -- Word 2013+ commentsExtended linkage -------------------------------

    @property
    def comments_extended_part(self) -> "CommentsExtendedPart | None":
        """Related |CommentsExtendedPart|, or |None| when none is related.

        Read-only view; does not create the part on demand. Use
        :meth:`comments_extended_part_or_add` to materialise one.

        .. versionadded:: 2026.05.10
        """
        try:
            return cast("CommentsExtendedPart", self.part_related_by(RT.COMMENTS_EXTENDED))
        except KeyError:
            return None

    def comments_extended_part_or_add(self) -> CommentsExtendedPart:
        """Return the related |CommentsExtendedPart|, creating one if needed.

        Materialises a default (empty) ``commentsExtended.xml`` part and
        relates it via ``RT.COMMENTS_EXTENDED`` on the first call.

        .. versionadded:: 2026.05.10
        """
        existing = self.comments_extended_part
        if existing is not None:
            return existing
        assert self.package is not None
        part = CommentsExtendedPart.default(self.package)
        self.relate_to(part, RT.COMMENTS_EXTENDED)
        return part
