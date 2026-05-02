.. _permissions:

Permissions and document protection
===================================

Word supports two complementary mechanisms for restricting edits:

- **Document protection** is a document-wide setting that locks the whole
  document into a mode — read-only, comments-only, tracked-changes, or
  forms-only — and optionally secures the lock with a password hash.
- **Permission ranges** carve out portions of a protected document that
  specific users or groups *are* allowed to edit. They are the escape
  hatch that makes "editable form in a locked template" workflows work.

|docx| surfaces both: document protection through
:attr:`.Settings.document_protection` and the convenience helpers
:meth:`.Settings.enable_protection`/:meth:`.Settings.disable_protection`,
and permission ranges through :meth:`.Paragraph.add_permission_range` and
:attr:`.Document.permission_ranges`.


Document protection
-------------------

Enabling protection
~~~~~~~~~~~~~~~~~~~

:meth:`.Settings.enable_protection` is the recommended entry point. It
creates the ``w:documentProtection`` element if absent, sets the mode,
enforces the restriction, and optionally hashes a password::

    >>> from docx import Document
    >>> from docx.enum.text import WD_PROTECTION
    >>> document = Document()
    >>> dp = document.settings.enable_protection(
    ...     WD_PROTECTION.COMMENTS,
    ...     password="s3cret",
    ...     enforce=True,
    ... )
    >>> dp.mode
    <WD_PROTECTION.COMMENTS: 1>
    >>> dp.enforce
    True

The supported modes are the members of :class:`.WD_PROTECTION`:

=================================  ====================  ===========================================
Member                             XML value             Behaviour
=================================  ====================  ===========================================
``WD_PROTECTION.READ_ONLY``        ``readOnly``          Document is read-only.
``WD_PROTECTION.COMMENTS``         ``comments``          Only comments may be added or modified.
``WD_PROTECTION.TRACKED_CHANGES``  ``trackedChanges``    All edits are recorded as tracked changes.
``WD_PROTECTION.FORMS``            ``forms``             Only form-field content may be edited.
=================================  ====================  ===========================================


Password hashing
~~~~~~~~~~~~~~~~

When a `password` is supplied, |docx| generates a random 16-byte salt and
hashes the password using Word's SHA-1 scheme with 100,000 iterations
(ISO/IEC 29500-1 §17.15.1.28). The resulting hash and salt are stored in
``@w:hash`` / ``@w:salt`` along with the algorithm metadata
(``cryptProviderType=rsaAES``, ``cryptAlgorithmSid=4``, ...).

Word's own implementation has historically had subtle variations across
versions; callers who need Word itself to accept the password at open time
should verify against their target Word release. For *detection* use cases
(reporting "this document is password-protected") the stored fields are
sufficient.


Disabling protection
~~~~~~~~~~~~~~~~~~~~

:meth:`.Settings.disable_protection` clears the mode and enforce flag but
leaves the ``w:documentProtection`` element in place so external tooling
that keyed off its presence still sees it::

    >>> document.settings.disable_protection()
    >>> document.settings.document_protection.mode is None
    True
    >>> document.settings.document_protection.enforce
    False


Fine-grained read access
~~~~~~~~~~~~~~~~~~~~~~~~

:attr:`.Settings.document_protection` returns a |DocumentProtection| proxy
exposing every underlying attribute individually: :attr:`mode`,
:attr:`enforce`, :attr:`formatting_locked`, :attr:`password_hash`,
:attr:`password_salt`, :attr:`crypto_provider_type`,
:attr:`crypto_algorithm_class`, :attr:`crypto_algorithm_type`,
:attr:`crypto_algorithm_sid`, and :attr:`spin_count`.


Permission ranges
-----------------

A permission range is delimited by ``w:permStart`` and ``w:permEnd`` markers
embedded in the body; between them the specified user or group may edit
even when the document is otherwise locked.

Adding a permission range
~~~~~~~~~~~~~~~~~~~~~~~~~

:meth:`.Paragraph.add_permission_range` wraps the calling paragraph in the
necessary markers::

    >>> p1 = document.add_paragraph("Editable by everyone.")
    >>> p2 = document.add_paragraph("Editable by Alice.")
    >>> p1.add_permission_range(edit_group="everyone")
    <docx.permissions.PermissionRange object at 0x...>
    >>> p2.add_permission_range(user="alice@example.com")

At least one of `edit_group` and `user` should typically be supplied — the
former for group restrictions (``"everyone"``, ``"current"``, or a named
group) and the latter for a single principal. The `name` parameter accepted
for API symmetry with :meth:`add_bookmark` is not persisted; ``w:permStart``
has no ``@w:name`` attribute.


Enumerating permission ranges
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

:attr:`.Document.permission_ranges` returns every permission range in the
document body in document order::

    >>> for pr in document.permission_ranges:
    ...     print(pr.id, pr.user, pr.edit_group)
    0 None everyone
    1 alice@example.com None

Each |PermissionRange| exposes:

- :attr:`.PermissionRange.id` — the integer identifier linking the
  matching ``permStart``/``permEnd`` pair.
- :attr:`.PermissionRange.user`, :attr:`.PermissionRange.edit_group`,
  :attr:`.PermissionRange.displaced_by_custom_xml` — the corresponding
  attributes on the underlying ``w:permStart``.


Deleting a permission range
~~~~~~~~~~~~~~~~~~~~~~~~~~~

:meth:`.PermissionRange.delete` removes both the start and end markers from
the document body::

    >>> document.permission_ranges[0].delete()

The body content between the markers is left untouched.


Scope and caveats
-----------------

- |docx| does not *enforce* the restrictions — that is Word's job at open
  time. Anything calling ``python-docx`` can freely modify every paragraph
  regardless of the protection mode, because the XML is just data to
  python-docx.
- Permission ranges added to the document body only cover that body.
  Ranges inside headers, footers, footnotes, or endnotes are not exposed
  via :attr:`.Document.permission_ranges`; reach for the paragraph-level
  :attr:`.Paragraph.permission_ranges` accessor in those containers.
- ``w:permStart`` IDs are assigned sequentially from zero; |docx| does not
  attempt to interleave them with custom IDs set by other tooling.
