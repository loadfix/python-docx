"""Custom element classes related to bookmarks and adjacent range markers.

This module models the ECMA-376 range-marker family that shares the
``CT_Markup`` / ``CT_MarkupRange`` / ``CT_Bookmark`` / ``CT_MoveBookmark``
inheritance chain from `wml.xsd`:

* ``w:bookmarkStart`` / ``w:bookmarkEnd`` — classic bookmarks.
* ``w:moveFromRangeStart`` / ``w:moveFromRangeEnd`` /
  ``w:moveToRangeStart`` / ``w:moveToRangeEnd`` — tracked-move markers.
* ``w:commentRangeStart`` / ``w:commentRangeEnd`` — comment range markers.

The shared ``@w:id`` attribute lives on every element (inherited from
``CT_Markup``); ``@w:name`` lives on the bookmark / move-bookmark variants;
``@w:author`` and ``@w:date`` live on the move variants (inherited from
``CT_TrackChange``).
"""

from __future__ import annotations

from docx.oxml.simpletypes import ST_DateTime, ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute


class CT_MarkupRange(BaseOxmlElement):
    """Base for every range-end marker (``CT_MarkupRange`` in wml.xsd).

    Used directly for ``w:bookmarkEnd``, ``w:moveFromRangeEnd``,
    ``w:moveToRangeEnd``, ``w:commentRangeStart`` and ``w:commentRangeEnd``.
    """

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]


class CT_BookmarkEnd(CT_MarkupRange):
    """`w:bookmarkEnd` element, marking the end of a bookmarked range.

    Schema: ``CT_MarkupRange`` — carries only ``@w:id``.
    """


class CT_BookmarkStart(BaseOxmlElement):
    """`w:bookmarkStart` element, marking the start of a bookmarked range.

    Schema: ``CT_Bookmark`` — extends ``CT_BookmarkRange`` with ``@w:name``.
    """

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    name: str = RequiredAttribute("w:name", ST_String)  # pyright: ignore[reportAssignmentType]


class CT_MoveBookmark(CT_BookmarkStart):
    """`w:moveFromRangeStart` / `w:moveToRangeStart`.

    Schema: ``CT_MoveBookmark`` — extends ``CT_Bookmark`` with the required
    ``@w:author`` and ``@w:date`` attributes inherited from ``CT_TrackChange``.
    ``@w:date`` is declared ``required`` on the move variant even though
    ``CT_TrackChange`` has it optional — we model it optional here so Word
    files that omit the date still parse.
    """

    author: str = RequiredAttribute("w:author", ST_String)  # pyright: ignore[reportAssignmentType]
    date: str | None = OptionalAttribute("w:date", ST_DateTime)  # pyright: ignore[reportAssignmentType]
