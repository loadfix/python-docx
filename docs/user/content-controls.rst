.. _content_controls:

Working with Content Controls
=============================

A *content control*, known in the OOXML schema as a *structured document tag* (SDT),
is a region of a Word document whose *kind of content* is declared up-front. You can
think of it as a typed placeholder. Word uses content controls to build form-like
documents where users can fill in answers, pick items from a list, or tick a box
without being able to edit the surrounding template text.

Each content control has a *type* (rich text, plain text, date, checkbox, combo box,
drop-down list, or picture), optional metadata (`tag`, `title`, and an integer `id`),
some *content* (runs or paragraphs held under a `w:sdtContent` element), and may
optionally be *data-bound* to an XML payload stored elsewhere in the package.

.. note::

   *python-docx* surfaces the structure and metadata of content controls, not Word's
   interactive behaviors. For example, a combo box's list of choices is carried
   through the XML untouched, but evaluating a data binding's XPath or enforcing
   editability locks is out of scope.


Content-control anatomy
-----------------------

In the XML, every content control looks like this::

    <w:sdt>
      <w:sdtPr>
        <w:alias w:val="Customer"/>        <!-- title -->
        <w:tag w:val="customer"/>          <!-- programmatic tag -->
        <w:id w:val="1234567890"/>         <!-- numeric id -->
        <!-- optional: type marker, e.g. <w:text/>, <w:date/>, <w14:checkbox/> -->
        <!-- optional: <w:dataBinding .../> -->
      </w:sdtPr>
      <w:sdtContent>
        <!-- runs (inline) or paragraphs (block-level) -->
      </w:sdtContent>
    </w:sdt>

- An SDT is **block-level** when it is a direct child of ``w:body`` or a table
  cell. Its `sdtContent` holds whole paragraphs.
- An SDT is **inline** when it is a child of a ``w:p``. Its `sdtContent` holds
  runs.
- The *type* is determined by a marker element inside ``w:sdtPr`` (``w:text``,
  ``w:date``, ``w:comboBox``, ``w:dropDownList``, ``w:picture``,
  ``w14:checkbox``). A rich-text SDT has no marker.


Creating content controls
-------------------------

Block-level content controls are added with :meth:`.Document.add_content_control`,
which appends a new `w:sdt` to the document body (just before any trailing section
properties)::

    >>> from docx import Document
    >>> from docx.content_controls import ContentControlType

    >>> document = Document()
    >>> cc = document.add_content_control(
    ...     ContentControlType.PLAIN_TEXT,
    ...     tag="customer",
    ...     title="Customer",
    ... )
    >>> cc
    <docx.content_controls.ContentControl object at 0x02468ACE>
    >>> cc.type
    <ContentControlType.PLAIN_TEXT: 'text'>
    >>> cc.tag, cc.title
    ('customer', 'Customer')
    >>> cc.sdt_id
    1872943104

Inline content controls are added via :meth:`.Paragraph.add_content_control`, which
appends the new `w:sdt` to that paragraph::

    >>> paragraph = document.add_paragraph("Hello, ")
    >>> inline = paragraph.add_content_control(
    ...     ContentControlType.RICH_TEXT, tag="greeting"
    ... )
    >>> inline.text = "world"

The full set of supported types is enumerated by :class:`.ContentControlType`:
``RICH_TEXT``, ``PLAIN_TEXT``, ``DATE``, ``CHECKBOX``, ``COMBO_BOX``, ``DROPDOWN``,
and ``PICTURE``. The rich-text flavor is the OOXML default and carries no explicit
marker element inside ``w:sdtPr``.

.. warning::

   The ``PICTURE`` type is surfaced for introspection only — *python-docx* does not
   yet provide a high-level API for assigning an image to a picture SDT.


Reading and modifying a content control
---------------------------------------

The :class:`.ContentControl` proxy exposes the metadata and content through simple
Python attributes::

    >>> cc.tag = "billing_customer"
    >>> cc.title = "Billing customer"
    >>> cc.type
    <ContentControlType.PLAIN_TEXT: 'text'>
    >>> cc.sdt_id  # read-only
    1872943104
    >>> cc.text = "Acme Co."
    >>> cc.text
    'Acme Co.'

Assigning to :attr:`.ContentControl.text` replaces the current `sdtContent` with a
single run (inline SDTs) or a single paragraph holding one run (block SDTs). To add
multiple paragraphs, runs with custom formatting, images, or tables, reach through
to the underlying XML via ``cc.element``.

Checkbox content controls carry an extra :attr:`.ContentControl.checked` property::

    >>> cbx = document.add_content_control(ContentControlType.CHECKBOX, tag="ok")
    >>> cbx.checked = True
    >>> cbx.checked
    True


Placeholder text
----------------

Word displays placeholder text inside an empty content control. In the XML this is
represented by a ``w:sdtPr/w:showingPlcHdr`` flag referencing a glossary document
entry. *python-docx* does not yet expose a first-class API for placeholder entries;
if you need to set a placeholder, you can do so by writing the initial text directly
into the control's content::

    >>> cc = document.add_content_control(ContentControlType.PLAIN_TEXT)
    >>> cc.text = "Click here to enter customer name"

When a user opens the document in Word and begins typing, the placeholder text will
be replaced with their input. This approach works for all supported types and is
the technique used by the behave fixtures that accompany this guide.


Iterating content controls
--------------------------

:attr:`.Document.content_controls` returns the block-level content controls found in
the main document story, in document order::

    >>> document.content_controls
    [<docx.content_controls.ContentControl object at 0x...>, ...]
    >>> for cc in document.content_controls:
    ...     print(cc.type, cc.tag, cc.title)

Only top-level `w:sdt` elements that are direct children of ``w:body`` are surfaced
here. Inline content controls appear under :attr:`.Paragraph.content_controls` on
the enclosing paragraph, which likewise yields controls in document order::

    >>> paragraph.content_controls
    [<docx.content_controls.ContentControl object at 0x...>]

Content controls nested inside table cells, headers, footers, or other stories are
not part of these collections. Walk the underlying XML tree (``.xpath(".//w:sdt")``
on the relevant element) to reach them if you need to.


Data binding
------------

A *data binding* ties the visible content of an SDT to an XPath expression over a
*custom XML data part* elsewhere in the package. Binding metadata is carried on the
control's ``w:sdtPr/w:dataBinding`` child::

    >>> cc = document.add_content_control(ContentControlType.PLAIN_TEXT, tag="customer")
    >>> binding = cc.set_data_binding(
    ...     xpath="/ns0:order[1]/ns0:customer[1]",
    ...     prefix_mappings="xmlns:ns0='http://example.com/orders'",
    ...     store_item_id="{11111111-2222-3333-4444-555555555555}",
    ... )
    >>> binding
    <docx.content_controls.DataBinding object at 0x02468ACE>
    >>> binding.xpath
    '/ns0:order[1]/ns0:customer[1]'
    >>> binding.prefix_mappings
    "xmlns:ns0='http://example.com/orders'"
    >>> binding.store_item_id
    '{11111111-2222-3333-4444-555555555555}'

Each attribute is read/write; reassigning :attr:`.DataBinding.xpath` or
:attr:`.DataBinding.store_item_id` updates the XML in place. A content control has
at most one data binding. Reading :attr:`.ContentControl.data_binding` on an unbound
control returns |None|. Use :meth:`.ContentControl.remove_data_binding` to clear it::

    >>> cc.remove_data_binding()
    >>> cc.data_binding is None
    True

.. note::

   *python-docx* does **not** evaluate the binding — it stores the XPath verbatim
   and leaves resolution to Word. If you need the bound value, fetch the
   corresponding :class:`.CustomXmlPart` and run the XPath yourself.


Custom XML data parts
---------------------

Data-bound content controls reference an XML payload stored in a sibling package
part: the *custom XML data part*. A typical document has one or more of these at
``/customXml/item{N}.xml``, each with a companion ``/customXml/itemProps{N}.xml``
part that declares a ``{GUID}``-formatted *store-item id* (and optional schema
references).

Those parts are surfaced read-only as :class:`.CustomXmlPart` proxies through
:attr:`.Document.custom_xml_parts`::

    >>> for part in document.custom_xml_parts:
    ...     print(part.partname, part.item_id, part.schema_refs)
    /customXml/item1.xml {EF278816-EC6F-A645-907D-7F25AECB1D4A} ['http://schemas.openxmlformats.org/officeDocument/2006/bibliography']
    /customXml/item2.xml {11111111-2222-3333-4444-555555555555} ['http://example.com/orders']

To resolve a binding to its backing part, match
:attr:`.DataBinding.store_item_id` to :attr:`.CustomXmlPart.item_id`::

    >>> target_id = cc.data_binding.store_item_id
    >>> part = next(p for p in document.custom_xml_parts if p.item_id == target_id)
    >>> part.root_element.tag
    '{http://example.com/orders}order'

Each proxy exposes :attr:`~docx.custom_xml.CustomXmlPart.blob` (raw bytes),
:attr:`~docx.custom_xml.CustomXmlPart.root_element` (parsed lxml element or |None|
on parse failure), :attr:`~docx.custom_xml.CustomXmlPart.item_id`, and
:attr:`~docx.custom_xml.CustomXmlPart.schema_refs`. The collection is read-only —
authoring new custom XML data parts is outside the scope of the current release and
requires working directly with the underlying OPC package.

See :ref:`content_controls_api` and :ref:`custom_xml_api` for the full API
reference.
