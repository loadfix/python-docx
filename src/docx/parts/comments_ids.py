"""Container parts for the Word 2016+ ``commentsIds.xml`` / 2018+
``commentsExtensible.xml`` auxiliary comment parts.

Modern Office writes two extra parts alongside the classic
``word/comments.xml``:

- ``word/commentsIds.xml`` (``w16cid:commentsIds``) — maps each legacy
  ``<w:comment>``'s integer ``@w:id`` to a stable paragraph id token
  (``@w16cid:paraId``) used by Office's threaded-reply feature.
- ``word/commentsExtensible.xml`` (``w16cex:commentsExtensible``) —
  attaches durable GUID-shaped identifiers (``@w16cex:durableId``) to
  each legacy comment so Office 365 clients don't renumber them across
  edit sessions.

Both parts are child relationships of ``word/comments.xml`` (target
sibling in ``word/``), matching the way Word writes them on save and
mirroring the ``commentsExtended.xml`` relationship convention.

.. versionadded:: 2026.05.10
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml.comments_ids import CT_CommentExtensibleList, CT_CommentIdList
from docx.oxml.ns import nsdecls
from docx.oxml.parser import parse_xml

if TYPE_CHECKING:
    from docx.package import Package


__all__ = ["CommentsExtensiblePart", "CommentsIdsPart"]


class CommentsIdsPart(XmlPart):
    """Container part for ``word/commentsIds.xml`` (``w16cid:commentsIds``).

    Root element is ``<w16cid:commentsIds>``; children are
    ``<w16cid:commentId>`` entries mapping a ``<w:comment>``'s
    ``@w:id`` to a ``@w16cid:paraId`` paragraph-id token.

    .. versionadded:: 2026.05.10
    """

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        element: CT_CommentIdList,
        package: "Package",
    ):
        super().__init__(partname, content_type, element, package)
        self._ids = element

    @property
    def element(self) -> CT_CommentIdList:
        """The root ``<w16cid:commentsIds>`` element of this part."""
        return self._ids

    @classmethod
    def default(cls, package: "Package") -> Self:
        """A newly created commentsIds part, containing an empty root."""
        partname = PackURI("/word/commentsIds.xml")
        content_type = CT.WML_COMMENTS_IDS
        xml = "<w16cid:commentsIds %s/>" % nsdecls("w16cid")
        element = cast("CT_CommentIdList", parse_xml(xml))
        return cls(partname, content_type, element, package)


class CommentsExtensiblePart(XmlPart):
    """Container part for ``word/commentsExtensible.xml`` (``w16cex:commentsExtensible``).

    Root element is ``<w16cex:commentsExtensible>``; children are
    ``<w16cex:commentExtensible>`` entries carrying a durable GUID-shaped
    identifier (``@w16cex:durableId``) for each legacy comment.

    .. versionadded:: 2026.05.10
    """

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        element: CT_CommentExtensibleList,
        package: "Package",
    ):
        super().__init__(partname, content_type, element, package)
        self._extensible = element

    @property
    def element(self) -> CT_CommentExtensibleList:
        """The root ``<w16cex:commentsExtensible>`` element of this part."""
        return self._extensible

    @classmethod
    def default(cls, package: "Package") -> Self:
        """A newly created commentsExtensible part, containing an empty root."""
        partname = PackURI("/word/commentsExtensible.xml")
        content_type = CT.WML_COMMENTS_EXTENSIBLE
        xml = "<w16cex:commentsExtensible %s/>" % nsdecls("w16cex")
        element = cast("CT_CommentExtensibleList", parse_xml(xml))
        return cls(partname, content_type, element, package)
