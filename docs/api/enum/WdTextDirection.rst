.. _WdTextDirection:

``WD_TEXT_DIRECTION``
=====================

Specifies the direction in which text flows within a table cell or section.

Example::

    from docx.enum.table import WD_TEXT_DIRECTION

    table = document.add_table(3, 3)
    table.cell(0, 0).text_direction = WD_TEXT_DIRECTION.TB_RL

----

LR_TB
    Left-to-right, top-to-bottom (default horizontal orientation).

TB_RL
    Top-to-bottom, right-to-left. Rotates text 90 degrees clockwise so it
    reads top-to-bottom along the right edge of the cell.

BT_LR
    Bottom-to-top, left-to-right. Rotates text 90 degrees counter-clockwise
    so it reads bottom-to-top along the left edge of the cell.

LR_TB_V
    Left-to-right horizontal flow with vertical glyph layout.

TB_RL_V
    Top-to-bottom, right-to-left vertical flow with vertical glyph layout.

TB_LR_V
    Top-to-bottom, left-to-right vertical flow with vertical glyph layout.
