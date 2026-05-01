"""Permission-range-related proxy types (`w:permStart` / `w:permEnd`)."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.oxml.permissions import CT_PermStart

if TYPE_CHECKING:
    from docx.oxml.document import CT_Body


class PermissionRange:
    """Proxy for a permission range defined by a `w:permStart`/`w:permEnd` pair.

    A permission range marks a portion of the document as editable by particular
    users or groups while the rest of the document is locked by document
    protection. See `ECMA-376 §17.13.2.17–18` for `w:permStart` / `w:permEnd`.
    """

    def __init__(self, permStart: CT_PermStart, body: CT_Body):
        self._permStart = permStart
        self._body = body

    @property
    def id(self) -> int:
        """The integer identifier linking this start to its matching `w:permEnd`."""
        return self._permStart.id

    @property
    def edit_group(self) -> str | None:
        """The group that may edit this range (`@w:edGrp`), or |None|.

        Common values include ``"everyone"`` and ``"current"``.
        """
        return self._permStart.edit_group

    @property
    def user(self) -> str | None:
        """The single user who may edit this range (`@w:ed`), or |None|."""
        return self._permStart.user

    @property
    def displaced_by_custom_xml(self) -> str | None:
        """Value of `@w:displacedByCustomXml`, or |None| when absent.

        Present when Word has displaced this range marker to a different spot
        because of a surrounding custom-XML element.
        """
        return self._permStart.displaced_by_custom_xml

    def delete(self) -> None:
        """Remove this permission range (start and matching end) from the document."""
        perm_id = str(self._permStart.id)
        # -- find and remove the matching permEnd --
        ends = self._body.xpath(f".//w:permEnd[@w:id='{perm_id}']")
        for end in ends:
            end.getparent().remove(end)
        # -- remove the permStart --
        self._permStart.getparent().remove(self._permStart)
