"""Element classes for the ``commentsExtended.xml`` (Word 2013+ ``w15:``) part.

This part stores the *resolved / reopened* state of comments and their
thread-reply parent linkage. Each ``<w15:commentEx>`` element references a
``<w:comment>`` (or a ``<w:p>`` within one) by paraId string. ``@w15:done``
captures whether the comment is marked resolved in the Word UI;
``@w15:parentParaId`` captures the reply-parent relationship (as a mirror
of the ``w16cid:paraIdParent`` already stored on ``<w:comment>`` itself).

``<w15:presenceInfo>`` is the co-authoring presence sidecar — its shape is
``{providerId, userId}``; Word writes it at the comments-extended root
when tracking the authors that are currently editing the document. We
model it as a first-class element so round-trip preserves the data.

This part is authored by Word when the user resolves a comment or creates
a threaded reply; it is always a child relationship of the
``word/comments.xml`` part (target is ``commentsExtended.xml`` in the same
``word/`` directory).

.. versionadded:: 2026.05.10
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.oxml.simpletypes import ST_OnOff, ST_String
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
)

if TYPE_CHECKING:
    pass


__all__ = [
    "CT_CommentExtended",
    "CT_CommentExtendedList",
    "CT_PresenceInfo",
]


# ---------------------------------------------------------------------------
# Root element
# ---------------------------------------------------------------------------


class CT_CommentExtendedList(BaseOxmlElement):
    """``<w15:commentsEx>`` — root of ``word/commentsExtended.xml``.

    Holds zero or more ``<w15:commentEx>`` children and (optionally) one or
    more ``<w15:presenceInfo>`` siblings tracking co-author presence.
    ``@w15:paraId`` on each ``<w15:commentEx>`` is the link key back to a
    paragraph inside a ``<w:comment>`` in ``word/comments.xml``.

    .. versionadded:: 2026.05.10
    """

    # -- type-declarations to fill in the gaps for metaclass-added methods --
    commentEx_lst: list["CT_CommentExtended"]
    presenceInfo_lst: list["CT_PresenceInfo"]

    commentEx = ZeroOrMore("w15:commentEx")
    presenceInfo = ZeroOrMore("w15:presenceInfo")

    _commentEx_xml_tmpl = (
        '<w15:commentEx {nsdecls}'
        ' w15:paraId="{para_id}" w15:done="{done}"'
        "/>"
    )

    _commentEx_with_parent_tmpl = (
        '<w15:commentEx {nsdecls}'
        ' w15:paraId="{para_id}" w15:done="{done}"'
        ' w15:paraIdParent="{parent_para_id}"'
        "/>"
    )

    @classmethod
    def new(cls) -> "CT_CommentExtendedList":
        """Return a new empty ``<w15:commentsEx>`` root element."""
        xml = '<w15:commentsEx %s/>' % nsdecls("w15")
        return cast("CT_CommentExtendedList", parse_xml(xml))

    def add_commentEx(
        self,
        paraId: str,
        done: bool = False,
        parentParaId: "str | None" = None,
    ) -> "CT_CommentExtended":
        """Append and return a new ``<w15:commentEx>`` for *paraId*.

        *done* sets ``@w15:done`` (the resolved-state flag). *parentParaId*
        sets ``@w15:paraIdParent`` when the comment is a threaded reply;
        omit or pass |None| for root-level comments.
        """
        done_str = "1" if done else "0"
        if parentParaId is None:
            xml = self._commentEx_xml_tmpl.format(
                nsdecls=nsdecls("w15"),
                para_id=paraId,
                done=done_str,
            )
        else:
            xml = self._commentEx_with_parent_tmpl.format(
                nsdecls=nsdecls("w15"),
                para_id=paraId,
                done=done_str,
                parent_para_id=parentParaId,
            )
        commentEx = cast("CT_CommentExtended", parse_xml(xml))
        self.append(commentEx)
        return commentEx

    def get_commentEx_by_paraId(
        self, paraId: str
    ) -> "CT_CommentExtended | None":
        """Return the ``<w15:commentEx>`` whose ``@w15:paraId`` matches *paraId*.

        Returns |None| when no matching element exists.
        """
        matches = cast(
            "list[CT_CommentExtended]",
            self.xpath(
                "./w15:commentEx[@w15:paraId=$paraId]",
                paraId=paraId,
            ),
        )
        return matches[0] if matches else None

    def get_children_for(self, paraId: str) -> "list[CT_CommentExtended]":
        """Return ``<w15:commentEx>`` elements whose ``@w15:paraIdParent`` matches *paraId*."""
        return cast(
            "list[CT_CommentExtended]",
            self.xpath(
                "./w15:commentEx[@w15:paraIdParent=$paraId]",
                paraId=paraId,
            ),
        )


# ---------------------------------------------------------------------------
# <w15:commentEx>
# ---------------------------------------------------------------------------


class CT_CommentExtended(BaseOxmlElement):
    """``<w15:commentEx>`` — extended metadata for a single comment.

    - ``@w15:paraId`` — link key; matches a paragraph paraId inside a
      ``<w:comment>`` in ``word/comments.xml`` (python-docx tracks this
      as ``w16cid:paraId`` on the ``<w:comment>`` itself — the paraId
      string is the same either way, the namespace difference is a
      Word extension history quirk).
    - ``@w15:done`` — ``xsd:boolean`` resolved/done flag.
    - ``@w15:paraIdParent`` — optional; reply-parent link.

    .. versionadded:: 2026.05.10
    """

    paraId: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w15:paraId", ST_String
    )
    done: bool = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w15:done", ST_OnOff
    )
    paraIdParent: "str | None" = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w15:paraIdParent", ST_String
    )


# ---------------------------------------------------------------------------
# <w15:presenceInfo>
# ---------------------------------------------------------------------------


class CT_PresenceInfo(BaseOxmlElement):
    """``<w15:presenceInfo>`` — co-authoring presence sidecar for an author.

    Word writes this inside ``commentsExtended.xml`` to track the
    provider / user identity pair for an author who currently has a
    presence-session open on the document. Read/write round-trips
    preserve the attribute pair; there is no parent/child linkage to
    individual comments (presence is a document-wide concern).

    - ``@w15:providerId`` — identity-provider name (e.g. ``"AD"`` for
      Active Directory, ``"Windows Live"`` for MSA).
    - ``@w15:userId`` — opaque user-identity string scoped to the
      provider.

    .. versionadded:: 2026.05.10
    """

    providerId: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w15:providerId", ST_String
    )
    userId: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w15:userId", ST_String
    )


# ---------------------------------------------------------------------------
# Namespace helpers used inside this module — kept here to avoid circular
# imports with ``docx.oxml.__init__``.
# ---------------------------------------------------------------------------


def _w15(tag: str) -> str:
    """Clark-name for a ``w15:<tag>`` QName."""
    return qn("w15:%s" % tag)
