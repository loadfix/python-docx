.. _WdProtection:

``WD_PROTECTION``
=================

Specifies the type of editing protection applied to a document.

Example::

    from docx.enum.text import WD_PROTECTION

    settings.document_protection.protection_type = WD_PROTECTION.READ_ONLY

----

READ_ONLY
    The document is read-only; no edits are permitted.

COMMENTS
    Only comments may be inserted or modified.

TRACKED_CHANGES
    Any edit is permitted, but is recorded as a tracked change.

FORMS
    Only form-field content may be edited.
