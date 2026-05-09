"""Collection providing access to comments added to this document."""

from __future__ import annotations

import datetime as dt
import secrets
from typing import TYPE_CHECKING, cast
from collections.abc import Iterator

from docx.blkcntnr import BlockItemContainer

if TYPE_CHECKING:
    from docx.oxml.comments import CT_Comment, CT_Comments
    from docx.oxml.comments_extended import (
        CT_CommentExtended,
        CT_CommentExtendedList,
    )
    from docx.parts.comments import CommentsPart
    from docx.parts.comments_extended import CommentsExtendedPart
    from docx.styles.style import ParagraphStyle
    from docx.text.paragraph import Paragraph


def _new_paragraph_id() -> str:
    """Return a fresh 8-character uppercase hex paraId (32-bit token).

    Used when auto-minting a ``w16cid:paraId`` entry for a new comment.
    Matches the token shape Word writes (``[0-9A-F]{8}``); uniqueness
    within the ``commentsIds.xml`` registry is enforced by the
    per-comment-id keying on ``set_paragraph_id``.
    """
    return secrets.token_hex(4).upper()


def _new_durable_id() -> str:
    """Return a fresh braced GUID token for a ``w16cex:durableId``.

    Word writes durable-ids in ``{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}``
    form (uppercase hex, surrounding braces). The value only has to be
    unique per-package; we use 16 random bytes and lay them out as a
    canonical UUID-like token without depending on the :mod:`uuid`
    module's v4 format details (which carry version bits we don't need
    to mirror for this use case).
    """
    b = secrets.token_hex(16).upper()
    return "{%s-%s-%s-%s-%s}" % (b[:8], b[8:12], b[12:16], b[16:20], b[20:])


class Comments:
    """Collection containing the comments added to this document."""

    def __init__(self, comments_elm: CT_Comments, comments_part: CommentsPart):
        self._comments_elm = comments_elm
        self._comments_part = comments_part

    def __iter__(self) -> Iterator[Comment]:
        """Iterator over the comments in this collection."""
        return (
            Comment(comment_elm, self._comments_part)
            for comment_elm in self._comments_elm.comment_lst
        )

    def __len__(self) -> int:
        """The number of comments in this collection."""
        return len(self._comments_elm.comment_lst)

    def add_comment(
        self,
        text: str = "",
        author: str = "",
        initials: str | None = "",
        date: dt.datetime | None = None,
    ) -> Comment:
        """Add a new comment to the document and return it.

        The comment is added to the end of the comments collection and is assigned a unique
        comment-id.

        If `text` is provided, it is added to the comment. This option provides for the common
        case where a comment contains a modest passage of plain text. Multiple paragraphs can be
        added using the `text` argument by separating their text with newlines (`"\\\\n"`).
        Between newlines, text is interpreted as it is in `Document.add_paragraph(text=...)`.

        The default is to place a single empty paragraph in the comment, which is the same
        behavior as the Word UI when you add a comment. New runs can be added to the first
        paragraph in the empty comment with `comments.paragraphs[0].add_run()` to adding more
        complex text with emphasis or images. Additional paragraphs can be added using
        `.add_paragraph()`.

        `author` is a required attribute, set to the empty string by default.

        `initials` is an optional attribute, set to the empty string by default. Passing |None|
        for the `initials` parameter causes that attribute to be omitted from the XML.

        `date` is the timestamp recorded on the comment's ``w:date`` attribute. When
        omitted, ``datetime.now(timezone.utc)`` is used. A timezone-aware datetime
        is honoured as-is (no conversion); a naive datetime is treated as already
        being in UTC. Pass |None| explicitly for the default behaviour.

        .. versionchanged:: 2026.05.0
           Added the `date` parameter so callers can supply a timezone-aware
           timestamp instead of the implicit ``datetime.now(UTC)``.
        """
        comment_elm = self._comments_elm.add_comment()
        comment_elm.author = author
        comment_elm.initials = initials
        comment_elm.date = date if date is not None else dt.datetime.now(dt.timezone.utc)
        comment = Comment(comment_elm, self._comments_part)

        # -- auto-mint w16cid:paraId + w16cex:durableId entries so Office 365
        # -- clients don't renumber them on the next edit. The inline
        # -- ``w:comment/@w16cid:paraId`` token is already allocated by
        # -- :meth:`CT_Comments.add_comment`; we mirror it into
        # -- ``word/commentsIds.xml`` and allocate a fresh GUID for
        # -- ``word/commentsExtensible.xml``. See R13-2 / ooxml-comments 0.4.
        comment.paragraph_id = comment_elm.paraId or _new_paragraph_id()
        comment.durable_id = _new_durable_id()

        if text == "":
            return comment

        para_text_iter = iter(text.split("\n"))

        first_para_text = next(para_text_iter)
        first_para = comment.paragraphs[0]
        first_para.add_run(first_para_text)

        for s in para_text_iter:
            comment.add_paragraph(text=s)

        return comment

    def get(self, comment_id: int) -> Comment | None:
        """Return the comment identified by `comment_id`, or |None| if not found."""
        comment_elm = self._comments_elm.get_comment_by_id(comment_id)
        return Comment(comment_elm, self._comments_part) if comment_elm is not None else None


class Comment(BlockItemContainer):
    """Proxy for a single comment in the document.

    Provides methods to access comment metadata such as author, initials, and date.

    A comment is also a block-item container, similar to a table cell, so it can contain both
    paragraphs and tables and its paragraphs can contain rich text, hyperlinks and images,
    although the common case is that a comment contains a single paragraph of plain text like a
    sentence or phrase.

    Note that certain content like tables may not be displayed in the Word comment sidebar due to
    space limitations. Such "over-sized" content can still be viewed in the review pane.
    """

    def __init__(self, comment_elm: CT_Comment, comments_part: CommentsPart):
        super().__init__(comment_elm, comments_part)
        self._comment_elm = comment_elm
        self._comments_part = comments_part

    def add_paragraph(self, text: str = "", style: str | ParagraphStyle | None = None) -> Paragraph:
        """Return paragraph newly added to the end of the content in this container.

        The paragraph has `text` in a single run if present, and is given paragraph style `style`.
        When `style` is |None| or ommitted, the "CommentText" paragraph style is applied, which is
        the default style for comments.
        """
        paragraph = super().add_paragraph(text, style)

        # -- have to assign style directly to element because `paragraph.style` raises when
        # -- a style is not present in the styles part
        if style is None:
            paragraph._p.style = "CommentText"  # pyright: ignore[reportPrivateUsage]

        return paragraph

    @property
    def author(self) -> str:
        """Read/write. The recorded author of this comment.

        This field is required but can be set to the empty string.
        """
        return self._comment_elm.author

    @author.setter
    def author(self, value: str):
        self._comment_elm.author = value

    @property
    def comment_id(self) -> int:
        """The unique identifier of this comment."""
        return self._comment_elm.id

    @property
    def initials(self) -> str | None:
        """Read/write. The recorded initials of the comment author.

        This attribute is optional in the XML, returns |None| if not set. Assigning |None| removes
        any existing initials from the XML.
        """
        return self._comment_elm.initials

    @initials.setter
    def initials(self, value: str | None):
        self._comment_elm.initials = value

    def add_reply(
        self,
        text: str = "",
        author: str = "",
        initials: str | None = "",
        date: dt.datetime | None = None,
    ) -> Comment:
        """Add a reply to this comment and return it.

        The reply is a new comment linked to this comment via the `w16cid:paraIdParent` attribute.
        Parameters behave identically to `Comments.add_comment()`, including the
        optional `date` timestamp (defaults to ``datetime.now(UTC)``).

        .. versionadded:: 2026.05.0
        .. versionchanged:: 2026.05.0
           Added the `date` parameter.
        """
        parent_para_id = self._comment_elm.paraId
        if parent_para_id is None:
            raise ValueError("Cannot add reply: parent comment has no paraId attribute.")

        comments_elm = cast("CT_Comments", self._comment_elm.getparent())
        reply_elm = comments_elm.add_reply(parent_para_id)
        reply_elm.author = author
        reply_elm.initials = initials
        reply_elm.date = date if date is not None else dt.datetime.now(dt.timezone.utc)
        reply = Comment(reply_elm, self._comments_part)

        # -- auto-mint commentsIds + commentsExtensible entries (R13-2). --
        reply.paragraph_id = reply_elm.paraId or _new_paragraph_id()
        reply.durable_id = _new_durable_id()

        if text == "":
            return reply

        para_text_iter = iter(text.split("\n"))

        first_para_text = next(para_text_iter)
        first_para = reply.paragraphs[0]
        first_para.add_run(first_para_text)

        for s in para_text_iter:
            reply.add_paragraph(text=s)

        return reply

    @property
    def replies(self) -> list[Comment]:
        """List of `Comment` objects that are replies to this comment.

        .. versionadded:: 2026.05.0
        """
        para_id = self._comment_elm.paraId
        if para_id is None:
            return []

        comments_elm = cast("CT_Comments", self._comment_elm.getparent())
        reply_elms = comments_elm.get_replies_for(para_id)
        return [Comment(reply_elm, self._comments_part) for reply_elm in reply_elms]

    def reply(
        self,
        text: str = "",
        author: str = "",
        initials: str | None = "",
        date: dt.datetime | None = None,
    ) -> Comment:
        """Add a threaded reply to this comment and return it.

        Alias for :meth:`add_reply` with a terser spelling. Identical
        behaviour — see :meth:`add_reply` for parameter semantics.

        .. versionadded:: 2026.05.10
        """
        return self.add_reply(text=text, author=author, initials=initials, date=date)

    # ------------------------------------------------------------------
    # Word 2013+ commentsExtended — resolved/done state + parent linkage
    # ------------------------------------------------------------------

    @property
    def is_resolved(self) -> bool:
        """|True| when this comment is marked resolved in Word (``w15:done``).

        A comment's resolved state lives in ``commentsExtended.xml`` as a
        ``<w15:commentEx>`` entry keyed on the comment's paraId. When no
        extended-comments part is related, or no ``<w15:commentEx>``
        exists for this comment's paraId, or ``@w15:done`` is omitted,
        this returns |False|.

        .. versionadded:: 2026.05.10
        """
        commentEx = self._commentEx
        if commentEx is None:
            return False
        return bool(commentEx.done)

    def resolve(self) -> None:
        """Mark this comment as resolved (``w15:done="1"``).

        Creates the ``commentsExtended.xml`` part and/or the
        ``<w15:commentEx>`` entry for this comment if they don't already
        exist. Raises :class:`ValueError` if the comment has no paraId
        (should never happen for python-docx-created comments, which
        always allocate one at insert time).

        .. versionadded:: 2026.05.10
        """
        self._set_done(True)

    def reopen(self) -> None:
        """Mark this comment as reopened (``w15:done="0"``).

        The inverse of :meth:`resolve`. Leaves the ``<w15:commentEx>``
        entry in place (Word does the same — it flips ``@w15:done`` to
        ``"0"`` rather than removing the entry) so any thread-reply
        linkage on the entry is preserved across reopen/resolve cycles.

        .. versionadded:: 2026.05.10
        """
        self._set_done(False)

    @property
    def parent_comment(self) -> "Comment | None":
        """The parent |Comment| this comment replies to, or |None|.

        Resolved via the ``<w:comment>``'s ``w16cid:paraIdParent``
        attribute (the primary link that python-docx authors on reply
        creation). Falls back to ``<w15:commentEx>/@w15:paraIdParent``
        when the inline attribute is absent (Word sometimes stores the
        parent only on the extended entry).

        .. versionadded:: 2026.05.10
        """
        parent_para_id = self._comment_elm.paraIdParent
        if parent_para_id is None:
            commentEx = self._commentEx
            if commentEx is not None:
                parent_para_id = commentEx.paraIdParent
        if parent_para_id is None:
            return None

        comments_elm = cast("CT_Comments", self._comment_elm.getparent())
        matches = comments_elm.xpath(
            "./w:comment[@w16cid:paraId=$paraId]",
            paraId=parent_para_id,
        )
        if not matches:
            return None
        return Comment(cast("CT_Comment", matches[0]), self._comments_part)

    # ------------------------------------------------------------------
    # Word 2016+ commentsIds.xml / Word 2018+ commentsExtensible.xml
    # ------------------------------------------------------------------

    @property
    def paragraph_id(self) -> str:
        """The stable paragraph id recorded in ``word/commentsIds.xml``.

        Returns the empty string when no entry exists for this comment's
        id yet. Read-through goes via the live ``<w16cid:commentsIds>``
        element on the related ``CommentsIdsPart`` — no registry is
        created on read-only access.

        Setting the attribute lazily creates ``word/commentsIds.xml``
        (and its relationship from ``word/comments.xml``) and writes or
        updates the entry mapping this comment's ``@w:id`` to *value*.

        .. versionadded:: 2026.05.10
        """
        part = self._comments_part.comments_ids_part
        if part is None:
            return ""
        entry = part.element.get_by_comment_id(self._comment_elm.id)
        return "" if entry is None else entry.paraId

    @paragraph_id.setter
    def paragraph_id(self, value: str) -> None:
        part = self._comments_part.comments_ids_part_or_add()
        part.element.set_paragraph_id(self._comment_elm.id, value)

    @property
    def durable_id(self) -> str:
        """The durable GUID-shaped id recorded in ``word/commentsExtensible.xml``.

        Office uses this identifier to re-attach comments across edit
        sessions without renumbering them. Matching is positional:
        entry *N* in ``commentsExtensible`` corresponds to the *N*-th
        ``<w16cid:commentId>`` entry in ``commentsIds`` (which is how
        Word writes the pair).

        Returns the empty string when no ``commentsIds`` entry exists
        for this comment (no durable-id slot to read) or when the
        ``commentsExtensible`` part has no matching positional entry.

        .. versionadded:: 2026.05.10
        """
        ids_part = self._comments_part.comments_ids_part
        ex_part = self._comments_part.comments_extensible_part
        if ids_part is None or ex_part is None:
            return ""
        ids = ids_part.element.iter_ids()
        for i, (cid, _) in enumerate(ids):
            if cid == self._comment_elm.id:
                entries = ex_part.element.commentExtensible_lst
                if i < len(entries):
                    return entries[i].durableId
                return ""
        return ""

    @durable_id.setter
    def durable_id(self, value: str) -> None:
        ids_part = self._comments_part.comments_ids_part_or_add()
        ex_part = self._comments_part.comments_extensible_part_or_add()
        ids = ids_part.element.iter_ids()
        entries = ex_part.element.commentExtensible_lst
        for i, (cid, _) in enumerate(ids):
            if cid == self._comment_elm.id:
                if i < len(entries):
                    entries[i].durableId = value
                    return
                # -- extend parallel entries so position N resolves --
                while len(ex_part.element.commentExtensible_lst) < i:
                    ex_part.element.set_durable_id(
                        "{00000000-0000-0000-0000-000000000000}"
                    )
                ex_part.element.set_durable_id(value)
                return
        raise LookupError(
            "cannot set durable_id: comment id %r has no commentsIds entry"
            % (self._comment_elm.id,)
        )

    # -- internal helpers ---------------------------------------------------

    @property
    def _commentEx(self) -> "CT_CommentExtended | None":
        """The ``<w15:commentEx>`` entry keyed on this comment's paraId.

        Returns |None| when no ``commentsExtended.xml`` part is related,
        or when no matching ``<w15:commentEx>`` exists yet.
        """
        para_id = self._comment_elm.paraId
        if para_id is None:
            return None
        ex_part = self._comments_part.comments_extended_part
        if ex_part is None:
            return None
        return ex_part.element.get_commentEx_by_paraId(para_id)

    def _set_done(self, value: bool) -> None:
        """Create or update the ``<w15:commentEx>/@w15:done`` for this comment."""
        para_id = self._comment_elm.paraId
        if para_id is None:
            raise ValueError(
                "cannot set resolved state: comment has no paraId "
                "(is this an imported legacy comment?)"
            )
        ex_part = self._comments_part.comments_extended_part_or_add()
        ex_root = ex_part.element
        commentEx = ex_root.get_commentEx_by_paraId(para_id)
        if commentEx is None:
            parent_para_id = self._comment_elm.paraIdParent
            ex_root.add_commentEx(
                paraId=para_id, done=value, parentParaId=parent_para_id
            )
        else:
            commentEx.done = value

    @property
    def text(self) -> str:
        """The text content of this comment as a string.

        Only content in paragraphs is included and of course all emphasis and styling is stripped.

        Paragraph boundaries are indicated with a newline (`"\\\\n"`)
        """
        return "\n".join(p.text for p in self.paragraphs)

    @property
    def timestamp(self) -> dt.datetime | None:
        """The date and time this comment was authored.

        This attribute is optional in the XML, returns |None| if not set.
        """
        return self._comment_elm.date
