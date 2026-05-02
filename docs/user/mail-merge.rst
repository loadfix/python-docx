.. _mail_merge:

Working with Mail Merge
=======================

Word supports *mail merge*, a feature in which a single *main document* is
combined with records drawn from an external *data source* to produce one
personalised output per record. Typical outputs include form letters, email
messages, envelopes, mailing labels, and faxes. *python-docx* does not
*execute* a mail merge, but it does expose the configuration block Word stores
inside ``word/settings.xml`` so that callers can read, construct, or remove it
programmatically.

When you open a main-merge document in Word, Word uses these stored settings
to know:

- what kind of merge to run (form letter, email, label, ...),
- where the merged output should go (a new document, a printer, email, ...),
- how to reach the data source (a connection string, an ODBC DSN, an Excel
  sheet, ...), and
- which rows to select from that data source (the stored query).

The configuration also records which record is currently "active" in Word's
preview, whether Word should display merged values instead of field
placeholders, and a handful of Boolean flags governing behaviour at merge
time.

**Scope.** *python-docx* surfaces the ``w:mailMerge`` element, its sub-elements,
and the three mail-merge enumerations. It does **not** create merge fields in
the document body, fetch data from external sources, or produce merged output
— the actual merge is still performed by Word (or by your own code reading the
configuration back out).


Accessing the mail-merge configuration
--------------------------------------

Every document exposes a :class:`.Settings` object via
:attr:`.Document.settings`. When the document has a ``w:mailMerge`` element,
:attr:`.Settings.mail_merge` returns a |MailMerge| proxy; when it does not, the
attribute is |None|::

    >>> from docx import Document
    >>> document = Document("contacts-form-letter.docx")
    >>> document.settings.mail_merge
    <docx.settings.MailMerge object at 0x02468ACE>

    >>> blank = Document()
    >>> blank.settings.mail_merge is None
    True


Enabling mail merge
-------------------

Use :meth:`.Settings.enable_mail_merge` to create (or replace) the
``w:mailMerge`` block. Every argument other than ``main_document_type`` is
optional; arguments left as |None| are simply omitted from the XML.

::

    >>> from docx import Document
    >>> from docx.enum.text import (
    ...     WD_MAIL_MERGE_DATA_TYPE,
    ...     WD_MAIL_MERGE_DESTINATION,
    ...     WD_MAIL_MERGE_TYPE,
    ... )
    >>> document = Document()
    >>> mail_merge = document.settings.enable_mail_merge(
    ...     main_document_type=WD_MAIL_MERGE_TYPE.EMAIL,
    ...     destination=WD_MAIL_MERGE_DESTINATION.EMAIL,
    ...     data_type=WD_MAIL_MERGE_DATA_TYPE.SPREADSHEET,
    ...     connect_string="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=contacts.xlsx",
    ...     query="SELECT FirstName, Email FROM [Sheet1$]",
    ...     mail_subject="Quarterly update",
    ...     address_field_name="Email",
    ... )

``enable_mail_merge()`` returns the |MailMerge| proxy so that additional
properties can be assigned in the same statement or on a subsequent line.
Calling it on a document that is already configured replaces the previous
``w:mailMerge`` element.

The simplest possible call produces a form-letter merge with no data source
attached::

    >>> document.settings.enable_mail_merge()


|MailMerge| properties
----------------------

Every |MailMerge| property is read/write and represents a single ``w:mailMerge``
child element. Assigning |None| (or |False| for the Boolean flags) removes the
underlying element.

.. rubric:: Typed, scalar properties

``main_document_type``
    A :ref:`WdMailMergeType` member identifying the merge kind. Reading returns
    |None| when the ``w:mainDocumentType`` child is absent.

``destination``
    A :ref:`WdMailMergeDestination` member describing where merged output is
    sent.

``data_type``
    A :ref:`WdMailMergeDataType` member describing the data-source kind.
    Unknown XML values read back as |None| rather than raising.

``connect_string``
    The raw connection string used to reach the data source (for example, an
    OLE DB or ODBC connection string). A plain |str| or |None|.

``query``
    The SQL-style query Word executes against the data source to select and
    order records. A plain |str| or |None|.

``mail_subject``
    The subject line used for email-destination merges.

``address_field_name``
    The name of the column inside the data source containing the recipient
    address (typically an email address column for email merges).

``active_record``
    The 1-based index of the record selected in Word's preview, as an |int|.
    Values that cannot be parsed as an integer read back as |None|.

``check_errors``
    Integer code controlling Word's error-reporting mode during merge.

.. rubric:: Boolean flags

The remaining properties correspond to on/off child elements. Each reads as
|True| when present and |False| when absent.

``link_to_query``
    Preserves the association between the stored query and the data source.

``do_not_suppress_blank_lines``
    Keeps blank output lines that would otherwise be suppressed when merge
    fields resolve to empty strings.

``mail_as_attachment``
    Sends the merged document as an email attachment rather than as the email
    body.

``view_merged_data``
    Tells Word to show merged field values rather than field placeholders when
    the document is opened.

Example of reading and updating properties::

    >>> mail_merge = document.settings.mail_merge
    >>> mail_merge.main_document_type
    <WD_MAIL_MERGE_TYPE.EMAIL: 4>
    >>> mail_merge.active_record
    3
    >>> mail_merge.view_merged_data
    True
    >>> mail_merge.mail_subject = "Updated subject"
    >>> mail_merge.mail_as_attachment = True


Disabling mail merge
--------------------

Call :meth:`.Settings.disable_mail_merge` to remove the ``w:mailMerge`` element
entirely. After the call, :attr:`.Settings.mail_merge` is |None|. The method is
idempotent — calling it on a document that has no ``w:mailMerge`` element is a
no-op::

    >>> document.settings.disable_mail_merge()
    >>> document.settings.mail_merge is None
    True


Mail-merge enumerations
-----------------------

Three enumerations live in :mod:`docx.enum.text`:

``WD_MAIL_MERGE_TYPE``
    Selects the main-document kind: ``CATALOG``, ``ENVELOPES``,
    ``MAILING_LABELS``, ``FORM_LETTERS`` (the default),
    ``EMAIL``, and ``FAX``.

``WD_MAIL_MERGE_DESTINATION``
    Selects the destination for the merged output: ``NEW_DOCUMENT``,
    ``PRINTER``, ``EMAIL``, and ``FAX``.

``WD_MAIL_MERGE_DATA_TYPE``
    Selects the data-source kind: ``TEXT_FILE``, ``DATABASE``,
    ``SPREADSHEET``, ``QUERY``, ``ODBC``, and ``NATIVE``.

See :ref:`WdMailMergeType`, :ref:`WdMailMergeDestination`, and
:ref:`WdMailMergeDataType` for the full enum reference.
