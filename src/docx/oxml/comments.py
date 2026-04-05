"""Custom element classes related to document comments."""

from __future__ import annotations

import datetime as dt
import secrets
from typing import TYPE_CHECKING, cast
from collections.abc import Callable

from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.oxml.simpletypes import ST_DateTime, ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrMore

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P

class CT_Comments(BaseOxmlElement):
    """`w:comments` element, the root element for the comments part.

    Simply contains a collection of `w:comment` elements, each representing a single comment. Each
    contained comment is identified by a unique `w:id` attribute, used to reference the comment
    from the document text. The offset of the comment in this collection is arbitrary; it is
    essentially a _set_ implemented as a list.
    """

    # -- type-declarations to fill in the gaps for metaclass-added methods --
    comment_lst: list[CT_Comment]

    comment = ZeroOrMore("w:comment")

    def add_comment(self) -> CT_Comment:
        """Return newly added `w:comment` child of this `w:comments`.

        The returned `w:comment` element is the minimum valid value, having a `w:id` value unique
        within the existing comments and the required `w:author` attribute present but set to the
        empty string. It's content is limited to a single run containing the necessary annotation
        reference but no text. Content is added by adding runs to this first paragraph and by
        adding additional paragraphs as needed.
        """
        next_id = self._next_available_comment_id()
        para_id = self._next_unique_para_id()
        comment = cast(CT_Comment, parse_xml(self._new_comment_xml(next_id, para_id)))
        self.append(comment)
        return comment

    def add_reply(self, parent_paraId: str) -> CT_Comment:
        """Return newly added reply `w:comment` child linked to the parent via `paraIdParent`.

        The reply comment has `w16cid:paraIdParent` set to `parent_paraId`, linking it to the
        parent comment.
        """
        next_id = self._next_available_comment_id()
        para_id = self._next_unique_para_id()
        comment = cast(CT_Comment, parse_xml(self._new_comment_xml(next_id, para_id)))
        comment.set(qn("w16cid:paraIdParent"), parent_paraId)
        self.append(comment)
        return comment

    def _new_comment_xml(self, comment_id: int, para_id: str) -> str:
        """Return XML string for a new `w:comment` element."""
        return (
            f'<w:comment {nsdecls("w", "w16cid")}'
            f' w:id="{comment_id}" w:author=""'
            f' w16cid:paraId="{para_id}">'
            f"  <w:p>"
            f"    <w:pPr>"
            f'      <w:pStyle w:val="CommentText"/>'
            f"    </w:pPr>"
            f"    <w:r>"
            f"      <w:rPr>"
            f'        <w:rStyle w:val="CommentReference"/>'
            f"      </w:rPr>"
            f"      <w:annotationRef/>"
            f"    </w:r>"
            f"  </w:p>"
            f"</w:comment>"
        )

    def get_replies_for(self, para_id: str) -> list[CT_Comment]:
        """Return list of `w:comment` elements whose `w16cid:paraIdParent` matches `para_id`."""
        return self.xpath(
            "./w:comment[@w16cid:paraIdParent=$paraId]",
            paraId=para_id,
        )

    def get_comment_by_id(self, comment_id: int) -> CT_Comment | None:
        """Return the `w:comment` element identified by `comment_id`, or |None| if not found."""
        comment_elms = self.xpath("(./w:comment[@w:id=$commentId])[1]", commentId=comment_id)
        return comment_elms[0] if comment_elms else None

    def _next_unique_para_id(self) -> str:
        """Generate a unique 8-character uppercase hex `paraId` value.

        The value is unique among all `w16cid:paraId` attributes in this `w:comments` element.
        """
        used_ids = set(self.xpath("./w:comment/@w16cid:paraId"))
        while True:
            para_id = secrets.token_hex(4).upper()
            if para_id not in used_ids:
                return para_id

    def _next_available_comment_id(self) -> int:
        """The next available comment id.

        According to the schema, this can be any positive integer, as big as you like, and the
        default mechanism is to use `max() + 1`. However, if that yields a value larger than will
        fit in a 32-bit signed integer, we take a more deliberate approach to use the first
        ununsed integer starting from 0.
        """
        used_ids = [int(x) for x in self.xpath("./w:comment/@w:id")]

        next_id = max(used_ids, default=-1) + 1

        if next_id <= 2**31 - 1:
            return next_id

        # -- fall-back to enumerating all used ids to find the first unused one --
        for expected, actual in enumerate(sorted(used_ids)):
            if expected != actual:
                return expected

        return len(used_ids)

class CT_Comment(BaseOxmlElement):
    """`w:comment` element, representing a single comment.

    A comment is a so-called "story" and can contain paragraphs and tables much like a table-cell.
    While probably most often used for a single sentence or phrase, a comment can contain rich
    content, including multiple rich-text paragraphs, hyperlinks, images, and tables.
    """

    # -- attributes on `w:comment` --
    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    author: str = RequiredAttribute("w:author", ST_String)  # pyright: ignore[reportAssignmentType]
    initials: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:initials", ST_String
    )
    date: dt.datetime | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:date", ST_DateTime
    )

    @property
    def paraId(self) -> str | None:
        """The `w16cid:paraId` attribute value, or |None| if not present."""
        return self.get(qn("w16cid:paraId"))

    @paraId.setter
    def paraId(self, value: str | None) -> None:
        if value is None:
            self.attrib.pop(qn("w16cid:paraId"), None)  # pyright: ignore[reportArgumentType]
        else:
            self.set(qn("w16cid:paraId"), value)

    @property
    def paraIdParent(self) -> str | None:
        """The `w16cid:paraIdParent` attribute value, or |None| if not present."""
        return self.get(qn("w16cid:paraIdParent"))

    @paraIdParent.setter
    def paraIdParent(self, value: str | None) -> None:
        if value is None:
            self.attrib.pop(qn("w16cid:paraIdParent"), None)  # pyright: ignore[reportArgumentType]
        else:
            self.set(qn("w16cid:paraIdParent"), value)

    # -- children --

    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    # -- type-declarations for methods added by metaclass --

    add_p: Callable[[], CT_P]
    p_lst: list[CT_P]
    tbl_lst: list[CT_Tbl]
    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    @property
    def inner_content_elements(self) -> list[CT_P | CT_Tbl]:
        """Generate all `w:p` and `w:tbl` elements in this comment."""
        return self.xpath("./w:p | ./w:tbl")
