.. _statistics:

Document statistics
===================

Word's **Review > Word Count** dialog summarizes the text content of a
document with four figures: *Pages*, *Words*, *Characters (no spaces)*, and
*Characters (with spaces)*, plus a *Paragraphs* count. |docx| provides an
equivalent summary through :attr:`.Document.statistics`, which returns a
|DocumentStatistics| named tuple.

Page counts are deliberately omitted because they depend on Word's pagination
engine, which does not run when a document is authored programmatically.
Everything else is computed directly from the body XML without requiring Word
to open the file.


Accessing the statistics
------------------------

::

    >>> from docx import Document
    >>> document = Document("report.docx")
    >>> stats = document.statistics
    >>> stats
    DocumentStatistics(paragraphs=42, words=3128, characters=19204,
        characters_no_spaces=16091)

The returned object is a |DocumentStatistics|, which is a
:class:`collections.namedtuple` subclass. Callers can destructure it directly::

    paragraphs, words, characters, characters_no_spaces = document.statistics

Or access fields by name::

    print(f"{document.statistics.words} words")


Field reference
---------------

``paragraphs``
    The count of non-empty body paragraphs. A paragraph is considered
    non-empty when it contains at least one non-whitespace character. This
    matches Word's behavior of excluding the "spacing" paragraphs that
    consist solely of whitespace or are entirely empty. The equivalent in
    Word's Word Count dialog is labeled **Paragraphs**.

``words``
    The count of whitespace-delimited tokens in the body text. A "word" is
    defined with :meth:`str.split` semantics — any run of non-whitespace
    characters surrounded by whitespace or string boundaries counts as one
    token. This corresponds to **Words** in Word's dialog.

``characters``
    The total count of characters in the body text, *including* spaces and
    other whitespace. This corresponds to **Characters (with spaces)** in
    Word's dialog.

``characters_no_spaces``
    The total count of characters in the body text, *excluding* any
    whitespace (spaces, tabs, and line breaks). This corresponds to
    **Characters (no spaces)** in Word's dialog.


Scope of the counts
-------------------

Only the main document story (the ``w:body``) is inspected. Text in headers,
footers, footnotes, endnotes, and comments is *not* included in the counts.
This matches Word's default "Word Count" behavior. Paragraphs nested inside
tables or block-level structured-document tags (content controls) *are*
included, because those paragraphs are part of the body story.

Because :attr:`.Document.statistics` is a read-only property, each access
recomputes the counts from the current state of the document. It is therefore
safe to call before and after content edits to observe how a change affects
the overall word or character count.


Building a simple report
------------------------

A typical use is producing a short summary for a pipeline log::

    stats = document.statistics
    print(
        f"{stats.paragraphs:>5}  paragraphs\n"
        f"{stats.words:>5}  words\n"
        f"{stats.characters:>5}  characters\n"
        f"{stats.characters_no_spaces:>5}  characters (no spaces)"
    )

Or enforcing an editorial policy — for instance, rejecting a submission that
falls below a minimum word count::

    MIN_WORDS = 500
    if document.statistics.words < MIN_WORDS:
        raise SystemExit(f"Document must contain at least {MIN_WORDS} words")

The underlying helper, :func:`docx.statistics.compute_statistics`, accepts a
``w:body`` element directly and is useful when you need the same counts for
something other than a top-level |Document|.
