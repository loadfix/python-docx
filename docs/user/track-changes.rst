.. _track_changes:

Working with Tracked Changes
============================

Word's *track changes* feature records every edit made to a document so that a
reviewer can later *accept* or *reject* each edit. When track-changes is on,
insertions, deletions, and formatting edits no longer silently modify the
document — they are wrapped in revision elements that carry the author's name,
the date and time of the edit, and (for moves) a pairing name that links the
source to the destination.

*python-docx* exposes the **read-side** of this model: collections of proxies
describing every tracked change in the document, plus a small preview helper
that renders paragraphs as plain strings with bracketed markers for inserted
and deleted runs. Accepting or rejecting changes is also available via
:meth:`Document.accept_all_changes` and :meth:`Document.reject_all_changes`;
finer-grained per-change accept/reject is planned for a future release.

.. note::
   Everything described below is read-only. *python-docx* does not yet expose
   authoring methods for creating tracked changes (e.g. a hypothetical
   ``Paragraph.add_tracked_insertion(...)``). Tracked-change content is
   expected to originate from Word itself (or from a workflow that edits the
   XML directly). If you need to synthesize track-change wrappers in tests or
   fixtures, drop down to the element level via ``lxml`` and the
   ``OxmlElement`` helpers used by the library internally.


Run-level tracked changes
-------------------------

Two run-level wrappers capture the most common edits:

* ``<w:ins>`` — one or more runs that were **inserted** by a reviewer
* ``<w:del>`` — one or more runs that were **deleted** (their text lives in
  ``<w:delText>`` rather than ``<w:t>``)

Every paragraph exposes these as |TrackedChange| objects via
:attr:`.Paragraph.tracked_changes`::

    >>> from docx import Document
    >>> document = Document("review-draft.docx")

    >>> paragraph = document.paragraphs[1]
    >>> paragraph.tracked_changes
    [<docx.tracked_changes.TrackedChange object at 0x02468ACE>,
     <docx.tracked_changes.TrackedChange object at 0x02468B12>]

Each proxy exposes the authorship metadata and the text of the change::

    >>> change = paragraph.tracked_changes[0]
    >>> change.type
    'deletion'
    >>> change.author
    'Bob'
    >>> change.date
    datetime.datetime(2025, 4, 10, 9, 0, tzinfo=datetime.timezone.utc)
    >>> change.text
    'brown'

The :attr:`.TrackedChange.type` property reports one of four string values:

* ``"insertion"`` — a ``<w:ins>`` wrapper
* ``"deletion"`` — a ``<w:del>`` wrapper
* ``"move_from"`` — the source side of a move revision (see below)
* ``"move_to"`` — the destination side of a move revision

To iterate every tracked change in the document body::

    >>> for paragraph in document.paragraphs:
    ...     for change in paragraph.tracked_changes:
    ...         print(f"{change.type:10s} {change.author:12s} {change.text!r}")
    deletion   Bob          'brown'
    insertion  Alice        'nimble'
    insertion  Carol        ', cruel world'

A paragraph with no tracked changes returns an empty list.


Move revisions
--------------

When a reviewer drags a selection of text from one paragraph to another with
track-changes on, Word records it as a **move revision**: the source is marked
``<w:moveFrom>`` (structurally a deletion whose text uses ``w:delText``) and
the destination is marked ``<w:moveTo>`` (structurally an insertion with
plain ``w:t``). Both wrappers carry a shared ``@w:name`` attribute pairing
them.

*python-docx* surfaces these as |MoveRevision|, a subclass of |TrackedChange|
that adds a ``name`` property and a ``peer`` lookup::

    >>> source_para = document.paragraphs[1]
    >>> move_from = source_para.tracked_changes[0]
    >>> type(move_from).__name__
    'MoveRevision'
    >>> move_from.type
    'move_from'
    >>> move_from.name
    'pair1'

    >>> peer = move_from.peer
    >>> peer.type
    'move_to'
    >>> peer.name
    'pair1'

``.peer`` walks up to the document root and searches the opposite side
(``w:moveTo`` when called on a ``w:moveFrom`` and vice versa) for the first
element whose ``@w:name`` matches. It returns |None| when the element has no
``@w:name`` or when no peer is found (unpaired halves can appear in
intermediate editing states).

.. note::
   Word also emits paragraph-level range markers
   ``<w:moveFromRangeStart/>``, ``<w:moveFromRangeEnd/>``,
   ``<w:moveToRangeStart/>``, and ``<w:moveToRangeEnd/>`` to bracket moves
   that span paragraph boundaries. These are *range* markers rather than
   run wrappers, so they are not exposed as |TrackedChange| proxies.
   They round-trip unchanged; callers needing to work with them can iterate
   the underlying XML.


Formatting changes
------------------

When a reviewer changes *formatting* — bold, alignment, section margins,
table or cell properties — Word records the prior state in a
*formatting-change* element rather than mutating the properties in place.
The revision is appended as a child of the relevant properties element:

=======================  ============================
Revision element         Parent element
=======================  ============================
``w:rPrChange``          ``w:rPr`` (on a run)
``w:pPrChange``          ``w:pPr`` (on a paragraph)
``w:sectPrChange``       ``w:sectPr`` (on a section)
``w:tblPrChange``        ``w:tblPr`` (on a table)
``w:tcPrChange``         ``w:tcPr`` (on a cell)
``w:trPrChange``         ``w:trPr`` (on a row)
=======================  ============================

Each revision element carries the same ``w:id`` / ``w:author`` / ``w:date``
metadata as the run-level wrappers, plus a single nested properties element
holding the *pre-revision* state. *python-docx* surfaces every one of them
through a |FormattingChange| proxy, accessible as a ``formatting_change``
property on the corresponding object:

* :attr:`.Run.formatting_change` — run formatting (``w:rPrChange``)
* :attr:`.Paragraph.formatting_change` — paragraph formatting (``w:pPrChange``)
* :attr:`.Section.formatting_change` — section formatting (``w:sectPrChange``)
* :attr:`.Table.formatting_change` — table-level formatting (``w:tblPrChange``)
* :attr:`._Cell.formatting_change` — cell formatting (``w:tcPrChange``)
* :attr:`._Row.formatting_change` — row formatting (``w:trPrChange``)

All six return |None| when the corresponding tracked revision is not present,
making the property easy to use as a predicate::

    >>> paragraph = document.paragraphs[1]
    >>> change = paragraph.formatting_change
    >>> change is None
    False
    >>> change.author
    'Bob'
    >>> change.date
    datetime.datetime(2025, 4, 10, 9, 5, tzinfo=datetime.timezone.utc)

The prior formatting is available via ``old_properties``::

    >>> old_pPr = change.old_properties
    >>> old_pPr.tag
    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr'

``old_properties`` returns the raw oxml element (``w:rPr``, ``w:pPr``,
``w:sectPr``, ``w:tblPr``, ``w:tcPr``, or ``w:trPr`` depending on the revision
kind). Callers that need to inspect specific properties can iterate its
children — there is no high-level proxy wrapping the historical state.


Table cell and row revisions
----------------------------

Word records **whole-cell** insertions and deletions via empty marker
elements inside the cell's ``w:tcPr``:

* ``<w:cellIns/>`` — the cell was inserted by a tracked change
* ``<w:cellDel/>`` — the cell was deleted by a tracked change

*python-docx* exposes these as boolean flags on |_Cell|::

    >>> table = document.tables[0]
    >>> cell = table.cell(0, 1)
    >>> cell.is_tracked_insertion
    True
    >>> cell.is_tracked_deletion
    False

    >>> deleted = table.cell(1, 0)
    >>> deleted.is_tracked_deletion
    True

Both flags are |False| when the cell has no ``w:tcPr`` or when neither
marker element is present. Row- and table-level property revisions use
``w:trPrChange`` and ``w:tblPrChange`` and are surfaced via the same
``formatting_change`` property described above.


revision_marks_text() preview
-----------------------------

For a quick terminal preview of tracked changes in running prose,
:meth:`.Paragraph.revision_marks_text` renders the paragraph as a plain
string with bracketed markers around inserted and deleted content::

    >>> paragraph = document.paragraphs[1]
    >>> paragraph.revision_marks_text()
    'The quick [-brown-][+nimble+] fox jumps.'

The default markers are CLI-friendly ``[+`` / ``+]`` for insertions and
``[-`` / ``-]`` for deletions. When the paragraph contains no tracked
changes the returned string matches :attr:`.Paragraph.text` exactly.

All four markers can be overridden::

    >>> paragraph.revision_marks_text(
    ...     open_ins="<INS>", close_ins="</INS>",
    ...     open_del="<DEL>", close_del="</DEL>",
    ... )
    'The quick <DEL>brown</DEL><INS>nimble</INS> fox jumps.'

Pass ANSI escape sequences for styled terminal output — for example
``"\033[32m"`` / ``"\033[0m"`` for green insertions and
``"\033[31m"`` / ``"\033[0m"`` for red deletions.

:meth:`.Document.revision_marks_text` calls the paragraph-level helper on
each top-level paragraph and joins the results with a blank-line separator
(``"\n\n"``). Tables in the body are skipped — this helper is meant as a
*quick preview* of prose, not a full-fidelity renderer::

    >>> print(document.revision_marks_text())
    Tracked insertions and deletions

    The quick [-brown-][+nimble+] fox jumps.

    Goodbye[+, cruel world+].

    Nothing to see here.


Revision-save IDs (rsid)
------------------------

Word stamps every paragraph and run with a **revision-save ID** — an
8-character hex string identifying the editing session during which the
element was last modified. The full set of session IDs lives in the
document settings under ``w:rsids``, with a single ``w:rsidRoot`` naming
the *first* session ever recorded for the document.

*python-docx* reads these values as plain strings.

Document-level ids are on |Settings|::

    >>> document.settings.rsid_root
    '00CAFE00'
    >>> document.settings.rsids
    ['00A1B2C3', '00DEAD00', '00BEEF00']

Per-element ids are on |Paragraph| and |Run|::

    >>> paragraph = document.paragraphs[1]
    >>> paragraph.rsid
    '00A1B2C3'

    >>> run = paragraph.runs[0]
    >>> run.rsid
    '00DEAD00'

Both return |None| when the element has no ``@w:rsidR`` attribute. RSIDs
are primarily useful to downstream *diff / merge* tooling: two runs with
the same RSID were last touched in the same editing session, so RSIDs
correlate edits across saves even when the text itself is unchanged. For a
stronger identifier that also accounts for position and content, see the
``stable_id`` property on |Run| and |Paragraph|.


Accepting or rejecting changes
------------------------------

Two document-level helpers apply every tracked revision in the body at
once::

    >>> n = document.accept_all_changes()
    >>> n
    5

    >>> n = document.reject_all_changes()

Accepting an insertion keeps the inserted content and removes the
``w:ins`` wrapper; accepting a deletion removes the wrapper *and* its
content. Rejecting does the opposite. Formatting revisions are resolved in
the same pass: accepting keeps the post-edit properties and discards the
change record; rejecting restores the pre-edit properties from
``old_properties``. Cell-level revisions (``w:cellIns``, ``w:cellDel``)
remove or preserve the enclosing cell as appropriate.

Both helpers return the count of change elements resolved.

.. note::
   Per-change ``TrackedChange.accept()`` and ``TrackedChange.reject()``
   methods and fine-grained filtering (by author, date range, or type) are
   planned for a future release. For now, prefer
   :meth:`.Document.accept_all_changes` /
   :meth:`.Document.reject_all_changes` when you need to resolve changes,
   and iterate |TrackedChange| objects when you need to inspect them.


What ends up in the XML
-----------------------

A run-level insertion looks approximately like:

.. code-block:: xml

    <w:p>
      <w:r><w:t xml:space="preserve">The quick </w:t></w:r>
      <w:ins w:id="2" w:author="Alice" w:date="2025-04-10T09:05:00Z">
        <w:r><w:t>nimble</w:t></w:r>
      </w:ins>
      <w:r><w:t xml:space="preserve"> fox jumps.</w:t></w:r>
    </w:p>

A deletion is structurally identical but uses ``<w:delText>`` in place of
``<w:t>``:

.. code-block:: xml

    <w:del w:id="1" w:author="Bob" w:date="2025-04-10T09:00:00Z">
      <w:r><w:delText>brown</w:delText></w:r>
    </w:del>

A move revision pairs two wrappers via ``@w:name``:

.. code-block:: xml

    <!-- source paragraph -->
    <w:moveFrom w:id="1" w:author="Alice" w:name="pair1" w:date="...">
      <w:r><w:delText>moved text</w:delText></w:r>
    </w:moveFrom>

    <!-- destination paragraph -->
    <w:moveTo w:id="2" w:author="Alice" w:name="pair1" w:date="...">
      <w:r><w:t>moved text</w:t></w:r>
    </w:moveTo>

A formatting revision nests the *old* properties element inside the change
wrapper:

.. code-block:: xml

    <w:pPr>
      <w:jc w:val="center"/>
      <w:pPrChange w:id="2" w:author="Bob" w:date="...">
        <w:pPr/>  <!-- old pPr: no w:jc, i.e. left-aligned -->
      </w:pPrChange>
    </w:pPr>


API reference
-------------

The tracked-changes proxies live in :mod:`docx.tracked_changes`; see
:ref:`tracked_changes_api` for the generated class documentation.
Relevant methods on |Document| are
:meth:`.Document.accept_all_changes`, :meth:`.Document.reject_all_changes`,
and :meth:`.Document.revision_marks_text`.
