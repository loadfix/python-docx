.. _WdMailMergeDestination:

``WD_MAIL_MERGE_DESTINATION``
=============================

Specifies the destination of mail-merge output.

Example::

    from docx.enum.text import WD_MAIL_MERGE_DESTINATION

    settings.mail_merge.destination = WD_MAIL_MERGE_DESTINATION.NEW_DOCUMENT

----

NEW_DOCUMENT
    Produce a new Word document containing the merged output.

PRINTER
    Send output directly to the printer.

EMAIL
    Email each merged record.

FAX
    Fax each merged record.
