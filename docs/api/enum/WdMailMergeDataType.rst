.. _WdMailMergeDataType:

``WD_MAIL_MERGE_DATA_TYPE``
===========================

Specifies the type of data source used by a mail-merge operation.

Example::

    from docx.enum.text import WD_MAIL_MERGE_DATA_TYPE

    settings.mail_merge.data_type = WD_MAIL_MERGE_DATA_TYPE.SPREADSHEET

----

TEXT_FILE
    Delimited text file (CSV / TSV).

DATABASE
    Microsoft Access or similar database.

SPREADSHEET
    Excel spreadsheet.

QUERY
    Word query file.

ODBC
    ODBC-connected data source.

NATIVE
    Native Word data source.
