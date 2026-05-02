.. _accessibility:

Checking document accessibility
===============================

Screen readers and other assistive technologies rely on a clean *heading
outline* to navigate a document. Word's own Accessibility Checker flags common
problems — skipped heading levels, multiple top-level headings, empty heading
paragraphs — and |docx| provides a small API that surfaces the same class of
issues so that build pipelines, CMS imports, and validation scripts can catch
them before a document is published.

The main entry point is :meth:`.Document.validate_heading_structure`, which
returns a list of |HeadingIssue| objects describing each problem it detects.
When the document's outline is clean the list is empty.


What counts as a heading
------------------------

A paragraph is considered a heading when its style name matches ``"Heading N"``
for ``N`` in ``1`` through ``9``, case-insensitively. Paragraphs with any other
style (including ``Title``, ``Subtitle``, or custom styles that merely *look*
like headings) are ignored by the validator. This matches how Word builds its
document outline from the built-in heading styles.


Detected issues
---------------

Each |HeadingIssue| carries three attributes:

``paragraph``
    The offending |Paragraph|, so callers can jump straight to the problem.

``kind``
    A short string identifier. One of:

    * ``"skipped_level"`` — a heading skips one or more outline levels
      (e.g. a ``Heading 3`` that directly follows a ``Heading 1`` without an
      intervening ``Heading 2``).
    * ``"multiple_h1"`` — the document contains more than one top-level
      heading. Only the *second* and later ``Heading 1`` paragraphs are
      flagged; the first is considered canonical.
    * ``"empty_heading"`` — a heading paragraph has no visible text after
      whitespace has been stripped.
    * ``"no_h1"`` — the first heading in the document is below
      ``Heading 1`` (e.g. the outline starts at ``Heading 2``).

``message``
    A human-readable description of the problem suitable for displaying to
    the author. These strings are not meant to be parsed; prefer branching on
    ``kind`` and building your own messages when you need i18n or tight UX
    control.


Running the validator
---------------------

Validation happens over the body's paragraphs (tables, headers, footers, and
comment text are not scanned). A minimal example::

    >>> from docx import Document
    >>> document = Document("report.docx")
    >>> issues = document.validate_heading_structure()
    >>> for issue in issues:
    ...     print(f"{issue.kind}: {issue.message}")
    skipped_level: Heading 3 follows Heading 1; Heading 2 is missing
    multiple_h1: Document contains more than one Heading 1; exactly one
        top-level heading is recommended

Issues are returned in document order, which makes them convenient to feed
straight into a linter-style report or to highlight inline in an editor.


Building an accessibility gate
------------------------------

A common pattern is to fail the build when *any* heading issues are present::

    issues = document.validate_heading_structure()
    if issues:
        for issue in issues:
            print(f"{issue.kind}: {issue.message}")
        raise SystemExit(1)

Or to selectively allow certain categories (for instance, tolerating multiple
``Heading 1`` paragraphs during a migration)::

    BLOCKING = {"skipped_level", "empty_heading", "no_h1"}
    blocking = [i for i in document.validate_heading_structure() if i.kind in BLOCKING]
    if blocking:
        raise SystemExit(1)


Working directly with the function
----------------------------------

:meth:`.Document.validate_heading_structure` is a thin wrapper around
:func:`docx.accessibility.validate_heading_structure`, which accepts *any*
iterable of |Paragraph| objects. This is handy for validating just one
section of a document, a comment, or a filtered view::

    from docx.accessibility import validate_heading_structure

    issues = validate_heading_structure(p for p in document.paragraphs if p.text)

The function is pure — it does not modify the paragraphs it inspects — so it
is safe to call repeatedly during a larger document-building workflow.
