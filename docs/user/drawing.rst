.. _drawing:

Working with DrawingML shapes
=============================

Word documents carry graphical content on a separate *drawing layer* alongside
the text layer. In addition to pictures, the drawing layer may host
*preset-geometry shapes* (rectangles, arrows, callouts, and similar), *group
shapes* that bundle multiple shapes together, *text frames* embedded in
shapes, *ink annotations* authored with a stylus, and *embedded OLE objects*
such as Excel workbooks or PDF files. For a conceptual introduction to the
two layers see :doc:`shapes`; this page is the fork-era companion and
documents the |docx| APIs for each of these drawing-layer features.

The drawing layer is expressed in OOXML as ``<w:drawing>`` elements nested
inside run elements. A single ``w:drawing`` may wrap an *inline* object
(``wp:inline`` — flows as a character glyph) or an *anchored / floating*
object (``wp:anchor`` — placed at arbitrary coordinates). The nested
``a:graphicData`` element carries one of ``pic:pic`` (a picture),
``c:chart`` (a chart reference), ``dgm:*`` (a SmartArt diagram),
``wpg:grpSp`` (a group of shapes), or ``wps:wsp`` (a DrawingML shape,
optionally with a ``wps:txbx`` text frame).

The :class:`docx.drawing.Drawing` proxy is the uniform entry point:
:attr:`.Drawing.type` returns a :class:`.WD_DRAWING_TYPE` member
(``PICTURE``, ``CHART``, ``DIAGRAM``, ``GROUP``, ``TEXT_BOX``, or
``SHAPE``), and the remaining methods on the proxy give you access to the
content-specific API.


Floating (anchored) images
--------------------------

A *floating image* is a picture wrapped in ``wp:anchor`` rather than
``wp:inline``. Unlike an inline picture — which behaves like a large glyph
— a floating image is positioned relative to a page, margin, column, or
paragraph, and surrounding text wraps around it according to a configurable
wrap style.

|docx| adds floating images via :meth:`.Paragraph.add_floating_image` and
exposes existing anchors via :attr:`.Paragraph.floating_images`. The proxy
type is :class:`docx.shape.FloatingImage`::

    >>> from docx import Document
    >>> from docx.enum.shape import WD_ANCHOR_H, WD_ANCHOR_V, WD_WRAP_TYPE
    >>> from docx.shared import Inches

    >>> document = Document()
    >>> paragraph = document.add_paragraph("Text surrounding the picture.")
    >>> floating = paragraph.add_floating_image(
    ...     "logo.png",
    ...     width=Inches(1.5),
    ...     position={
    ...         "h_anchor": WD_ANCHOR_H.PAGE,
    ...         "v_anchor": WD_ANCHOR_V.PAGE,
    ...         "horizontal": Inches(2),
    ...         "vertical": Inches(3),
    ...         "wrap": WD_WRAP_TYPE.SQUARE,
    ...     },
    ... )
    >>> floating.horizontal_anchor, floating.vertical_anchor
    (<WD_ANCHOR_H.PAGE: 'page'>, <WD_ANCHOR_V.PAGE: 'page'>)
    >>> floating.horizontal_offset, floating.vertical_offset
    (1828800, 2743200)

The ``position`` dict is optional; when omitted the image is anchored at
``COLUMN``/``PARAGRAPH`` with :class:`.WD_WRAP_TYPE.SQUARE` and zero
offsets. Only the keys you supply are overridden; unspecified keys fall back
to the default. The :attr:`.FloatingImage.position` property returns a dict
in the same shape that you can feed back into a subsequent call on a
different anchor.

Each horizontal/vertical anchor choice corresponds to a
``wp:positionH/@relativeFrom`` or ``wp:positionV/@relativeFrom`` token from
the OOXML grammar. The ``wrap`` entry maps onto ``wp:wrapSquare``,
``wp:wrapTight``, ``wp:wrapThrough``, ``wp:wrapTopAndBottom``, or
``wp:wrapNone`` (``BEHIND``/``IN_FRONT`` are both ``wp:wrapNone`` with
different ``behindDoc`` attributes).

Floating images are enumerated per paragraph::

    >>> for paragraph in document.paragraphs:
    ...     for fi in paragraph.floating_images:
    ...         print(fi.wrap_type, fi.offset)


Preset-geometry shapes
----------------------

|docx| can add DrawingML preset shapes — the kind of geometric primitives
you reach via ``Insert > Shapes`` in Word — inline to a paragraph::

    >>> from docx.enum.shape import WD_SHAPE
    >>> from docx.shared import Inches

    >>> paragraph = document.add_paragraph()
    >>> shape = paragraph.add_shape(
    ...     WD_SHAPE.ROUNDED_RECTANGLE,
    ...     width=Inches(2),
    ...     height=Inches(1),
    ...     text="Click me",
    ... )
    >>> shape.shape_type
    <WD_SHAPE.ROUNDED_RECTANGLE: 'roundRect'>
    >>> shape.text
    'Click me'

:meth:`.Paragraph.add_shape` returns a :class:`docx.drawing.WordprocessingShape`
proxy wrapping the newly-created ``wps:wsp`` element. The proxy exposes
:attr:`.WordprocessingShape.name`, :attr:`.WordprocessingShape.shape_type`
(a member of :class:`.WD_SHAPE`), and a read/write
:attr:`.WordprocessingShape.text` property.

The implemented :class:`.WD_SHAPE` members cover rectangles
(``RECTANGLE`` / ``ROUNDED_RECTANGLE``), ovals (``OVAL``), arrows
(``ARROW_RIGHT``), and a rounded-rectangle callout
(``CALLOUT_ROUNDED_RECTANGLE``). Shapes authored in Word with other preset
geometries round-trip correctly: a read via
:attr:`.Paragraph.drawings` reports them as :class:`.WD_DRAWING_TYPE.SHAPE`,
and :attr:`.WordprocessingShape.shape_type` returns |None| when the preset
token does not correspond to a known enum member.

``add_shape`` validates its ``shape_type`` argument::

    >>> paragraph.add_shape("rect")
    Traceback (most recent call last):
      ...
    TypeError: shape_type must be a WD_SHAPE member, got 'rect'


Group shapes
------------

Word can combine several shapes into a *group* — a single unit that can be
selected, moved, or resized as a whole. The underlying element is
``wpg:grpSp`` and groups may nest arbitrarily. |docx| models groups read-only
through :class:`docx.drawing.GroupShape`::

    >>> from docx.drawing import GroupShape, WordprocessingShape

    >>> for paragraph in document.paragraphs:
    ...     for drawing in paragraph.drawings:
    ...         if drawing.is_group:
    ...             group = drawing.group_shape
    ...             print(group.name)
    ...             for child in group.shapes:
    ...                 print("  ", type(child).__name__)

Each child returned by :attr:`.GroupShape.shapes` is a
:class:`.WordprocessingShape`, a :class:`docx.drawing.Picture`, or a nested
:class:`.GroupShape`. Unsupported child element types (for example
``wpg:graphicFrame``) are omitted from the list so calling code can assume
every entry is one of the three proxy classes.

Use :attr:`.Drawing.group_shapes` to access every top-level group on a
drawing; :attr:`.Drawing.group_shape` returns just the first (which is what
Word writes for a single selection) or |None| if the drawing is not a group.


Text box (shape text-frame) content
-----------------------------------

A ``wps:wsp`` shape may carry a text frame (``wps:txbx/w:txbxContent``) —
this is what Word exposes as "Edit Text" on a shape. When a shape contains
text its :attr:`.Drawing.type` reports :class:`.WD_DRAWING_TYPE.TEXT_BOX`
instead of ``SHAPE``.

Two access paths are available. :attr:`.Drawing.text` returns a single
concatenated string (multiple paragraphs are separated by ``\n``), and
:attr:`.Drawing.paragraphs` returns the |Paragraph| objects inside the text
frame so the full run-level API is available::

    >>> drawing = paragraph.drawings[0]
    >>> drawing.type
    <WD_DRAWING_TYPE.TEXT_BOX: 2>
    >>> drawing.text
    'First line\nSecond line\nThird line'
    >>> [p.text for p in drawing.paragraphs]
    ['First line', 'Second line', 'Third line']

:meth:`.Paragraph.add_shape` accepts an optional ``text`` argument; when
supplied, a minimal text frame containing that string is attached to the new
``wps:wsp``. :attr:`.WordprocessingShape.text` is read/write — assigning a
string replaces the existing text-frame content::

    >>> shape = paragraph.add_shape(WD_SHAPE.RECTANGLE, text="Initial")
    >>> shape.text
    'Initial'
    >>> shape.text = "Replaced"
    >>> shape.text
    'Replaced'


Ink annotations
---------------

Word on touch-enabled devices can record stylus-drawn *ink annotations*.
They are stored as separate ``word/ink/ink*.xml`` parts in the
`InkML <http://www.w3.org/2003/InkML>`_ format and referenced from the
document body by a ``<w:contentPart r:id="..."/>`` element inside a run.

|docx| exposes ink annotations read-only via
:attr:`.Document.ink_annotations` and :attr:`.Paragraph.ink_annotations`.
The proxy type is :class:`docx.ink.InkAnnotation`::

    >>> for annotation in document.ink_annotations:
    ...     print(annotation.partname, annotation.stroke_count)
    /word/ink/ink1.xml 2
    /word/ink/ink2.xml 1

:attr:`.InkAnnotation.blob` returns the raw InkML XML bytes so you can pass
them to a downstream parser or renderer. :attr:`.InkAnnotation.stroke_count`
reports the number of ``inkml:trace`` elements in the part — this is a
structural count, not a glyph count; it counts both direct children of
``inkml:ink`` and traces nested inside ``inkml:traceGroup``.

``w:contentPart`` references whose relationship target is missing from the
package, or whose target part is of the wrong type, are silently skipped
rather than raised. This keeps repair-mode loads of damaged documents from
crashing when a stray ink reference was left behind after a part was
dropped. python-docx does not support *creating* or *modifying* ink
annotations; the API is deliberately read-only.


Embedded OLE objects
--------------------

Word supports embedding OLE objects — Excel workbooks, PDF documents,
mathematical equations, and so on — directly into a document. Each object is
stored as a separate part (usually under ``word/embeddings/``) whose content
type is ``application/vnd.openxmlformats-officedocument.oleObject``. The
reference comes from an ``<o:OLEObject>`` element inside a ``<w:object>``
element inside a run.

|docx| exposes embedded objects read-only via
:attr:`.Document.embedded_objects` and
:attr:`.Paragraph.embedded_objects`. The proxy type is
:class:`docx.embedded_objects.EmbeddedObject`::

    >>> for obj in document.embedded_objects:
    ...     print(obj.prog_id, obj.type, len(obj.blob))
    Excel.Sheet.12 Embed 16

:attr:`.EmbeddedObject.prog_id` is the ProgID token identifying the object's
type (``Excel.Sheet.12``, ``AcroExch.Document``, ``Equation.DSMT4``, etc.).
:attr:`.EmbeddedObject.type` is either ``"Embed"`` (the binary lives in the
package) or ``"Link"`` (the binary lives at a file-system or URL target).
:attr:`.EmbeddedObject.blob` returns the raw OLE bytes.

A reference whose relationship id cannot be resolved — for example because
the target part was dropped or is of the wrong type — still produces an
:class:`.EmbeddedObject`, but its :attr:`.EmbeddedObject.blob` returns
``b""`` and its :attr:`.EmbeddedObject.embedded_partname` returns |None|.
Callers that care can filter on ``if obj.blob:``. Creation and modification
are intentionally not supported.


Accessibility: alt text and titles
----------------------------------

Every inline picture, floating picture, preset shape, and group in a Word
document has an accessibility-facing *description* (alt text) and an
optional *title*. These map onto the ``@descr`` and ``@title`` attributes of
the ``wp:docPr`` element inside the ``wp:inline`` or ``wp:anchor``. For
assistive technologies the description is read in place of the image when
the text layer is dictated aloud.

|docx| exposes both attributes as read/write properties on
:class:`.InlineShape` and :class:`.FloatingImage`::

    >>> shape = document.inline_shapes[0]
    >>> shape.alt_text = "A pencil-drawing of a mountain peak"
    >>> shape.title = "Mountain peak"
    >>> shape.alt_text
    'A pencil-drawing of a mountain peak'

Either attribute can be assigned |None| to clear it. When the underlying
XML attribute is absent the getter returns |None|; for floating images
whose ``wp:docPr`` element itself is absent the getter still returns
|None| and the setter creates the element on demand.

Setting alt text is the single most effective accessibility fix available
for a document containing graphical content; aim to populate
``alt_text`` on every decorative or informational picture in a document you
generate.


SVG pictures
------------

|docx| accepts SVG (Scalable Vector Graphics) files as input to
:meth:`.Run.add_picture`. Word renders SVG natively in recent versions; for
compatibility with older consumers the library also stores a small PNG
*fallback* so the image is visible even when the reader does not understand
SVG.

Both are referenced from the same ``pic:pic`` element: the fallback PNG is
the primary ``a:blip/@r:embed`` and the SVG is attached via an
``asvg:svgBlip`` extension element. Round-tripping an SVG preserves the
original bytes — the fallback is generated once at write time.

::

    >>> paragraph = document.add_paragraph()
    >>> run = paragraph.add_run()
    >>> run.add_picture("diagram.svg", width=Inches(3))

SVG pixel dimensions are inferred from the ``width``/``height`` or
``viewBox`` attributes on the root ``<svg>`` element. Unitless values are
treated as CSS pixels; ``in`` / ``cm`` / ``mm`` / ``pt`` units are converted
to pixels at 96 DPI (the CSS reference density). An SVG whose dimensions
cannot be parsed falls back to the SVG spec default of 300 x 150.

Floating placement (``add_floating_image``) does not implement the PNG
fallback path and treats SVG like any other image — it relies on the
consumer to render SVG directly. If you need SVG on the drawing layer with
wide-compatibility fallback, add it inline.


Iterating drawings generically
------------------------------

:attr:`.Paragraph.drawings` returns a :class:`.Drawing` proxy for every
``w:drawing`` descendant of the paragraph, regardless of what the drawing
wraps. Use :attr:`.Drawing.type` to branch on the content kind::

    >>> from docx.enum.shape import WD_DRAWING_TYPE

    >>> for paragraph in document.paragraphs:
    ...     for drawing in paragraph.drawings:
    ...         if drawing.type is WD_DRAWING_TYPE.PICTURE:
    ...             image = drawing.image
    ...             ...
    ...         elif drawing.type is WD_DRAWING_TYPE.CHART:
    ...             chart = drawing.chart
    ...             ...
    ...         elif drawing.type is WD_DRAWING_TYPE.GROUP:
    ...             group = drawing.group_shape
    ...             ...
    ...         elif drawing.type is WD_DRAWING_TYPE.TEXT_BOX:
    ...             print(drawing.text)
    ...         elif drawing.type is WD_DRAWING_TYPE.SHAPE:
    ...             # bare wps:wsp with no text frame
    ...             ...

The dedicated collections — :attr:`.Document.inline_shapes`,
:attr:`.Document.charts`, :attr:`.Document.ink_annotations`,
:attr:`.Document.embedded_objects` — are the right tool for single-kind
surveys; :attr:`.Paragraph.drawings` is the right tool when you need
position-aware (paragraph-scoped) enumeration or when you want to handle
every drawing kind in one pass.
