.. _WdMailMergeType:

``WD_MAIL_MERGE_TYPE``
======================

Specifies the type of mail-merge operation.

Example::

    from docx.enum.text import WD_MAIL_MERGE_TYPE

    settings.mail_merge.merge_type = WD_MAIL_MERGE_TYPE.FORM_LETTERS

----

CATALOG
    Catalog-style merge (all records on one page).

ENVELOPES
    Envelope printing merge.

MAILING_LABELS
    Mailing-label printing merge.

FORM_LETTERS
    Form-letter merge (one letter per record).

EMAIL
    Email-message merge.

FAX
    Fax merge.
