.. _watermarks:

Page watermarks
===============

Word supports two kinds of watermark — text and image — both rendered in the
*default page header* as VML shapes. |docx| exposes them per-section through
four members on :class:`.Section`:

- :meth:`.Section.add_text_watermark` — attach a text watermark (``"DRAFT"``,
  ``"CONFIDENTIAL"``, etc.).
- :meth:`.Section.add_image_watermark` — attach an image watermark
  (company logo, classification stamp, ...).
- :attr:`.Section.watermark` — read the current watermark back as a
  |Watermark| proxy, or |None| if none is present.
- :meth:`.Section.remove_watermark` — clear any existing watermark.


Why the default header?
-----------------------

Word represents watermarks as VML ``v:shape`` elements living inside a header
paragraph. Painting them in the header is how they end up on every page of
the section without being part of the body text. |docx| follows the same
convention: adding a watermark detaches the section's header from the
previous section (so section-specific watermarks remain independent) and
paints the shape into the default header.

As a consequence, if a section's header is linked to the previous section,
calling :meth:`add_text_watermark` or :meth:`add_image_watermark`
automatically sets ``is_linked_to_previous = False`` on that section's
:attr:`.Section.header` before adding the shape.


Adding a text watermark
-----------------------

::

    >>> from docx import Document
    >>> from docx.shared import Pt, RGBColor
    >>> document = Document()
    >>> section = document.sections[0]
    >>> watermark = section.add_text_watermark(
    ...     text="DRAFT",
    ...     font="Arial",
    ...     size=Pt(80),
    ...     color=RGBColor(0x80, 0x80, 0x80),
    ...     layout="diagonal",
    ... )
    >>> watermark.type, watermark.text
    ('text', 'DRAFT')

The `font`, `size`, `color`, and `layout` parameters are optional. Omitting
them yields Word's standard draft-watermark look — 72-point Calibri in
silver (``#C0C0C0``) drawn diagonally across the page.

Only ``"diagonal"`` and ``"horizontal"`` are accepted for `layout`. Any
other value raises :class:`ValueError`.


Adding an image watermark
-------------------------

::

    >>> from docx.shared import Inches
    >>> watermark = section.add_image_watermark(
    ...     "logo.png",
    ...     width=Inches(3),
    ...     height=Inches(2),
    ... )
    >>> watermark.type
    'image'
    >>> watermark.text is None
    True

`image_path` may be a filesystem path or a file-like object (any stream
:class:`docx.image.Image` accepts). When both `width` and `height` are
omitted the image's native dimensions are used; otherwise the image is
scaled to the supplied dimensions while preserving aspect ratio.

Either call replaces any watermark already present — there is no
``add_additional_watermark`` variant because Word itself does not support
stacked watermarks in a single section.


Reading the current watermark
-----------------------------

:attr:`.Section.watermark` returns the |Watermark| proxy for the section's
default header, or |None| when:

- the header is still linked to the previous section, or
- the header does not contain a watermark shape.

::

    >>> document.sections[0].watermark
    <docx.watermark.Watermark object at 0x...>
    >>> document.sections[0].watermark.type
    'text'


Removing a watermark
--------------------

:meth:`.Section.remove_watermark` strips every watermark paragraph from the
section's default header. It is safe to call on a section that has no
watermark; the method is a no-op in that case::

    >>> section.remove_watermark()
    >>> section.watermark is None
    True


Scope and limitations
---------------------

- Watermarks are per-section. A document with multiple sections can have a
  different watermark in each; use :meth:`add_text_watermark` /
  :meth:`add_image_watermark` on each |Section| individually.
- |docx| writes VML watermark shapes — this matches what Word has always
  done for watermarks, including Word 365. Other word processors' support
  for VML varies; LibreOffice in particular renders VML watermarks with
  reduced fidelity.
- Colour is specified as RGB; Word's "washout" / semi-transparent effect is
  not separately exposed, but supplying a light grey such as ``RGBColor(0xC0,
  0xC0, 0xC0)`` approximates it.
- Font rotation beyond the ``diagonal`` / ``horizontal`` presets is not
  parameterised; callers needing more control can modify the returned shape
  element directly.
