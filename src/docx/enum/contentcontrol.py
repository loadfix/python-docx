"""Enumerations related to content controls (structured document tags)."""

from __future__ import annotations

from docx.enum.base import BaseEnum


class WD_CONTENT_CONTROL_TYPE(BaseEnum):
    """Specifies the type of a content control (structured document tag).

    Example::

        from docx.enum.contentcontrol import WD_CONTENT_CONTROL_TYPE

        content_control = document.content_controls[0]
        if content_control.type == WD_CONTENT_CONTROL_TYPE.PLAIN_TEXT:
            print(content_control.text)
    """

    RICH_TEXT = (0, "Rich text content control.")
    """Rich text content control (default when no specific type element is present)."""

    PLAIN_TEXT = (1, "Plain text content control.")
    """Plain text content control."""

    CHECKBOX = (2, "Checkbox content control.")
    """Checkbox content control."""

    COMBO_BOX = (3, "Combo box content control.")
    """Combo box content control."""

    DROP_DOWN = (4, "Drop-down list content control.")
    """Drop-down list content control."""

    DATE = (5, "Date picker content control.")
    """Date picker content control."""

    PICTURE = (6, "Picture content control.")
    """Picture content control."""
