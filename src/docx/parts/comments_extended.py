"""Container part for the Word 2013+ ``commentsExtended.xml`` part.

The extended-comments part carries per-comment resolved/done state and
the threaded-reply parent linkage, keyed off ``@w15:paraId`` → the
``w16cid:paraId`` (or ``w14:paraId``) string on a comment paragraph in
``word/comments.xml``.

The part relationship is typically emitted by Word from the
``word/comments.xml`` part rather than from the document body part; we
follow that convention so the produced package matches what Word
writes on save.

.. versionadded:: 2026.05.10
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml.comments_extended import CT_CommentExtendedList
from docx.oxml.ns import nsdecls
from docx.oxml.parser import parse_xml

if TYPE_CHECKING:
    from docx.package import Package


__all__ = ["CommentsExtendedPart"]


class CommentsExtendedPart(XmlPart):
    """Container part for ``word/commentsExtended.xml`` (Word 2013+ ``w15:``).

    Root element is ``<w15:commentsEx>``; children are ``<w15:commentEx>``
    (per-comment resolved/done state and thread-reply parent link) and
    (optionally) ``<w15:presenceInfo>`` (co-author presence sidecar).
    """

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        element: CT_CommentExtendedList,
        package: "Package",
    ):
        super().__init__(partname, content_type, element, package)
        self._comments_ex = element

    @property
    def element(self) -> CT_CommentExtendedList:
        """The root ``<w15:commentsEx>`` element of this part."""
        return self._comments_ex

    @classmethod
    def default(cls, package: "Package") -> Self:
        """A newly created commentsExtended part, containing an empty root."""
        partname = PackURI("/word/commentsExtended.xml")
        content_type = CT.WML_COMMENTS_EXTENDED
        xml = '<w15:commentsEx %s/>' % nsdecls("w15", "mc")
        element = cast("CT_CommentExtendedList", parse_xml(xml))
        return cls(partname, content_type, element, package)
