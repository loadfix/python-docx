"""Enumerations related to structured document tags (content controls)."""

from __future__ import annotations

import enum


class WD_CONTENT_CONTROL_TYPE(enum.Enum):
    """Content control types for structured document tags.

    Example::

        from docx.enum.sdt import WD_CONTENT_CONTROL_TYPE

        cc = document.content_controls[0]
        if cc.type == WD_CONTENT_CONTROL_TYPE.CHECKBOX:
            print("It's a checkbox!")
    """

    PLAIN_TEXT = "plainText"
    """Plain-text content control."""

    RICH_TEXT = "richText"
    """Rich-text content control."""

    CHECKBOX = "checkbox"
    """Checkbox content control."""

    COMBO_BOX = "comboBox"
    """Combo box content control."""

    DROP_DOWN = "dropDown"
    """Drop-down list content control."""

    DATE = "date"
    """Date picker content control."""

    PICTURE = "picture"
    """Picture content control."""
