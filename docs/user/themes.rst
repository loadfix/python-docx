.. _themes:

Themes
======

Every Word document ships with a **theme** — a coordinated set of colours
and fonts that style elements refer to rather than hard-coding values. The
theme lives in ``word/theme/theme1.xml`` and is referenced indirectly from
every paragraph and character style. Changing the theme therefore
re-skins the whole document without touching individual paragraphs.

|docx| exposes the theme **read-only**: :attr:`.Document.theme` returns a
|Theme| proxy when the document has a theme relationship, or |None| when it
doesn't (hand-authored packages can legally omit the theme part).

Theme authoring is not supported. If you need a different theme, design it
in Word and use that document as your template.


Retrieving the theme
--------------------

::

    >>> from docx import Document
    >>> document = Document()
    >>> theme = document.theme
    >>> theme.name
    'Office Theme'
    >>> theme.colors.name
    'Office'
    >>> theme.fonts.name
    'Office'

:attr:`.Theme.name`, :attr:`.ThemeColors.name`, and :attr:`.ThemeFonts.name`
return the human-readable names assigned to the theme as a whole and to its
two schemes; they are all nullable strings.


Colour scheme
-------------

Word's colour scheme has twelve slots named with short OOXML tokens. |docx|
exposes them both by token and by accessor:

``dk1``
    :attr:`.ThemeColors.dark_1` — dark body text (usually black).

``lt1``
    :attr:`.ThemeColors.light_1` — light background (usually white).

``dk2``
    :attr:`.ThemeColors.dark_2` — secondary dark colour.

``lt2``
    :attr:`.ThemeColors.light_2` — secondary light colour.

``accent1`` … ``accent6``
    :attr:`.ThemeColors.accent_1` through
    :attr:`.ThemeColors.accent_6` — additional accents used by charts,
    tables, and shape fills.

``hlink``
    :attr:`.ThemeColors.hyperlink` — unvisited hyperlink.

``folHlink``
    :attr:`.ThemeColors.followed_hyperlink` — visited hyperlink.

Every accessor returns an |RGBColor| or |None| (when the slot is missing or
its value cannot be resolved to RGB — for example, an ``a:sysClr`` without
a ``lastClr`` fallback)::

    >>> theme.colors.dark_1
    RGBColor(0x00, 0x00, 0x00)
    >>> theme.colors.accent_1
    RGBColor(0x4F, 0x81, 0xBD)
    >>> theme.colors.hyperlink
    RGBColor(0x00, 0x00, 0xFF)

The subscript form uses the OOXML token directly — handy when the slot name
is coming from data rather than a hard-coded attribute::

    >>> theme.colors["accent1"]
    RGBColor(0x4F, 0x81, 0xBD)
    >>> theme.colors["hlink"]
    RGBColor(0x00, 0x00, 0xFF)
    >>> theme.colors["bogus"]
    Traceback (most recent call last):
      ...
    KeyError: 'bogus'

A |None| value means the slot is defined *but* its child element cannot be
resolved to an RGB triple; a :class:`KeyError` means the token is not one
of the twelve legal slot names. This lets callers distinguish "slot omitted"
from "typo in slot name".


Font scheme
-----------

The font scheme pairs two typeface bundles:

- **Major** fonts, used for headings (``a:majorFont``);
- **Minor** fonts, used for body text (``a:minorFont``).

Each bundle nests three slots for the primary script regions:

====================  ===========================================  ===================================
Slot                  Accessor                                     OOXML element
====================  ===========================================  ===================================
Major Latin           :attr:`.ThemeFonts.major_latin`              ``a:majorFont/a:latin/@typeface``
Minor Latin           :attr:`.ThemeFonts.minor_latin`              ``a:minorFont/a:latin/@typeface``
Major East Asian      :attr:`.ThemeFonts.major_east_asian`         ``a:majorFont/a:ea/@typeface``
Minor East Asian      :attr:`.ThemeFonts.minor_east_asian`         ``a:minorFont/a:ea/@typeface``
Major Complex Script  :attr:`.ThemeFonts.major_cs`                 ``a:majorFont/a:cs/@typeface``
Minor Complex Script  :attr:`.ThemeFonts.minor_cs`                 ``a:minorFont/a:cs/@typeface``
====================  ===========================================  ===================================

Each accessor returns the typeface string, or |None| when the slot is
missing::

    >>> theme.fonts.major_latin
    'Calibri'
    >>> theme.fonts.minor_latin
    'Cambria'


When theme is |None|
--------------------

A document built from hand-authored parts — or a document that has had its
theme deliberately stripped — will cause :attr:`.Document.theme` to return
|None|. Guard callers accordingly::

    >>> theme = document.theme
    >>> if theme is None:
    ...     print("no theme")
    ... else:
    ...     print(theme.colors.accent_1)


Interoperability
----------------

- Styles reference the theme indirectly through ``w:themeColor`` on
  ``w:color`` and ``w:themeFont`` on ``w:rFonts``. Changing theme colours in
  Word updates every style that uses them automatically; changing an
  individual style's colour overrides the theme reference for that style
  only.
- Some Word features (theme-aware tables, for example) use theme tint and
  shade attributes to derive new colours from the twelve base slots. |docx|
  exposes only the base slots; tinted/shaded variants must be computed by
  the caller.
