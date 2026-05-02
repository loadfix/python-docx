.. _WdTableAutofit:

``WD_TABLE_AUTOFIT``
====================

Specifies the autofit behavior for a table.

Example::

    from docx.enum.table import WD_TABLE_AUTOFIT

    table = document.add_table(3, 3)
    table.autofit_behavior = WD_TABLE_AUTOFIT.AUTOFIT_TO_CONTENTS

----

AUTOFIT_TO_WINDOW
    Column widths adjust automatically so the table fills the window width.

AUTOFIT_TO_CONTENTS
    Column widths adjust automatically based on cell contents.

FIXED_WIDTH
    Column widths are fixed regardless of cell contents.
