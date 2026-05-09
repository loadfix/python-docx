"""Re-export of ``w16cid:`` / ``w16cex:`` commentsIds element classes.

The canonical implementations live in :mod:`ooxml_comments.oxml.commentsids`;
this module exposes them under the historical ``docx.oxml.comments_ids``
import path so downstream consumers continue to work.

- ``<w16cid:commentsIds>`` / ``<w16cid:commentId>`` — modern Office's
  paragraph-id registry for legacy comments (``word/commentsIds.xml``).
- ``<w16cex:commentsExtensible>`` / ``<w16cex:commentExtensible>`` — modern
  Office's durable-GUID registry for legacy comments
  (``word/commentsExtensible.xml``).

Neither part is in ECMA-376 proper (both are Microsoft 2016 / 2018
extensions) but Word writes them on every recent docx with comments;
without preserving them, Office 365 will renumber paragraph-ids and
durable-ids on the next save.

.. versionadded:: 2026.05.10
"""

from __future__ import annotations

from ooxml_comments.oxml.commentsids import (
    CT_CommentExtensible,
    CT_CommentExtensibleList,
    CT_CommentId,
    CT_CommentIdList,
)

__all__ = [
    "CT_CommentExtensible",
    "CT_CommentExtensibleList",
    "CT_CommentId",
    "CT_CommentIdList",
]
