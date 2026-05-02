.. _form_fields:

Working with Legacy Form Fields
===============================

Word supports two families of in-document form controls. The modern family,
*content controls* (Structured Document Tags, or SDTs), was introduced in
Word 2007. The older family — *legacy form fields* — predates them and is
still widely used, especially by templates authored for older Word versions
and by documents produced by legal, accounting, and government tooling.

*python-docx* exposes the legacy family via the ``docx.form_fields`` module.

**Legacy Form-Field Anatomy.** A legacy form field is a *complex field* whose
``begin`` ``w:fldChar`` carries a ``w:ffData`` child. The ``w:ffData``
element holds the form field's metadata (name, help text, enabled flag,
calc-on-exit flag) and a type-specific options block:

- ``w:textInput`` for a *text-input* field (``FORMTEXT``) — a free-form
  text entry with an optional default value, max length, and format.
- ``w:checkBox`` for a *checkbox* field (``FORMCHECKBOX``) — a boolean
  state with a default and a current ``checked`` state.
- ``w:ddList`` for a *dropdown* field (``FORMDROPDOWN``) — a list of
  options with a default-selection index and a result-selection index.

Each form field is bracketed by the usual complex-field markers: a ``begin``
fldChar, an ``instrText`` run (``FORMTEXT``, ``FORMCHECKBOX``, or
``FORMDROPDOWN``), a ``separate`` fldChar, a *result* region (runs carrying
the rendered value), and an ``end`` fldChar.

**Form-Field Name.** Each form field has a *name* (``w:ffData/w:name``) that
acts as the programmatic identifier used by Word VBA macros and by ``REF``
fields elsewhere in the document to retrieve the field's value. Names do not
need to be unique, but Word tooling typically treats them as such.

**Read vs. Mutate.** The :class:`FormField` proxy is read-oriented: it
exposes a type discriminator, the shared metadata (name, help text, status
text, enabled, calc-on-exit), a ``value`` derivation, and per-type views
(:class:`TextInputFormField`, :class:`CheckboxFormField`,
:class:`DropdownFormField`) that are *read-only* projections over the
corresponding ``w:ffData`` child.

To *create* new form fields, three paragraph-level convenience methods are
provided — :meth:`Paragraph.add_text_form_field`,
:meth:`Paragraph.add_checkbox_form_field`, and
:meth:`Paragraph.add_dropdown_form_field` — each of which appends a
complete complex-field sequence to the paragraph and returns a
:class:`FormField` proxy.

**Applicability.** Legacy form fields render and round-trip correctly in
Word. They can appear in the document body, in table cells, and in headers
and footers. The :attr:`Document.form_fields` collection walks *top-level
body paragraphs only* — to access form fields nested inside table cells,
headers, footers, footnotes, or endnotes, iterate the enclosing
paragraphs and read their :attr:`Paragraph.form_fields` collections.


Accessing the form-fields collection
------------------------------------

The top-level collection is accessed via :attr:`Document.form_fields`::

    >>> from docx import Document
    >>> document = Document("application-form.docx")
    >>> fields = document.form_fields
    >>> len(fields)
    3
    >>> [ff.name for ff in fields]
    ['FullName', 'Subscribe', 'Country']

Each member is a :class:`FormField` proxy. The
:attr:`FormField.type` property returns a :class:`WD_FORM_FIELD_TYPE`
enum member that discriminates the three field families::

    >>> from docx.form_fields import WD_FORM_FIELD_TYPE
    >>> fields[0].type
    <WD_FORM_FIELD_TYPE.TEXT: 'text'>
    >>> fields[0].type is WD_FORM_FIELD_TYPE.TEXT
    True

The shared metadata is exposed on the proxy itself::

    >>> ff = fields[0]
    >>> ff.name
    'FullName'
    >>> ff.help_text
    ''
    >>> ff.enabled
    True
    >>> ff.calc_on_exit
    False


Text-input form fields
----------------------

Type-specific attributes are exposed via a narrow read-only view. For text
inputs, use :attr:`FormField.text_input`::

    >>> ff = fields[0]  # -- the FullName text input --
    >>> ti = ff.text_input
    >>> ti.default
    'Jane Doe'
    >>> ti.max_length
    40
    >>> ti.format
    ''

A ``max_length`` of |None| indicates no limit (the ``w:maxLength`` element
is absent or its value is ``0``, the OOXML "no-limit" sentinel).

The current rendered value is available on the proxy itself via
:attr:`FormField.value`, which returns the concatenated text of the runs
between the ``separate`` and ``end`` markers::

    >>> ff.value
    'Jane Doe'


Checkbox form fields
--------------------

Checkbox views are exposed via :attr:`FormField.checkbox`::

    >>> ff = fields[1]  # -- the Subscribe checkbox --
    >>> cb = ff.checkbox
    >>> cb.default
    True
    >>> cb.checked
    True

When ``w:checked`` is absent but ``w:default`` is present, ``checked``
returns the default — mirroring Word's runtime behaviour. For checkboxes,
:attr:`FormField.value` returns the boolean ``checked`` state directly::

    >>> ff.value
    True


Dropdown form fields
--------------------

Dropdown views are exposed via :attr:`FormField.dropdown`::

    >>> ff = fields[2]  # -- the Country dropdown --
    >>> dd = ff.dropdown
    >>> dd.options
    ['US', 'UK', 'AU']
    >>> dd.default_index
    1
    >>> dd.result_index
    1

``default_index`` and ``result_index`` are 0-based. When ``w:result`` is
absent, ``result_index`` falls back to ``default_index``. For dropdowns,
:attr:`FormField.value` returns the *selected option string* (the entry at
``result_index``), or the empty string when the index is out of range::

    >>> ff.value
    'UK'


Adding form fields
------------------

Form fields are appended to a paragraph via three convenience methods on
:class:`Paragraph`. Each returns a :class:`FormField` proxy for the newly
added field::

    >>> from docx import Document
    >>> document = Document()

    >>> p = document.add_paragraph("Name: ")
    >>> p.add_text_form_field(name="FullName", default="Jane Doe", maxlength=40)
    <docx.form_fields.FormField object at 0x02468ACE>

    >>> p = document.add_paragraph("Subscribe? ")
    >>> p.add_checkbox_form_field(name="Subscribe", checked=True)
    <docx.form_fields.FormField object at 0x02468ACE>

    >>> p = document.add_paragraph("Country: ")
    >>> p.add_dropdown_form_field(
    ...     name="Country", options=["US", "UK", "AU"], default_index=1,
    ... )
    <docx.form_fields.FormField object at 0x02468ACE>

    >>> document.save("application-form.docx")

Each method emits a complete complex-field sequence — the ``begin`` run
(with the ``w:ffData`` attached to its ``w:fldChar``), the ``instrText``
run, the ``separate`` run, a *result* run, and the ``end`` run. The
rendered result text is seeded so Word displays the initial value
immediately without a field update.

For the type-specific attributes of these methods — such as ``maxlength``
for text inputs, or ``default_index`` for dropdowns — see the API
documentation at :ref:`form_fields_api`.
