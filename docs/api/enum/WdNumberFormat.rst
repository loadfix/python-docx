.. _WdNumberFormat:

``WD_NUMBER_FORMAT``
====================

Specifies a numeric format used for numbering list items, footnotes, or endnotes.

Example::

    from docx.enum.text import WD_NUMBER_FORMAT

    level.number_format = WD_NUMBER_FORMAT.UPPER_ROMAN

----

DECIMAL
    Decimal numbers (1, 2, 3 ...).

ARABIC
    Alias for ``DECIMAL`` (Arabic numerals: 1, 2, 3 ...).

UPPER_ROMAN
    Uppercase Roman numerals (I, II, III ...).

LOWER_ROMAN
    Lowercase Roman numerals (i, ii, iii ...).

UPPER_LETTER
    Uppercase letters (A, B, C ...).

LOWER_LETTER
    Lowercase letters (a, b, c ...).

ORDINAL
    Ordinal numbers (1st, 2nd, 3rd ...).

CARDINAL_TEXT
    Cardinal text (One, Two, Three ...).

ORDINAL_TEXT
    Ordinal text (First, Second ...).

CHICAGO
    Chicago Manual of Style footnote marks (``*``, dagger, double dagger, section).

BULLET
    Bullet character (not numbered).

NONE
    No number.
