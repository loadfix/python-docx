"""Layer 3: Round-trip test helpers.

Create a document, save it, re-open it, and assert the data reads back correctly.
"""

from __future__ import annotations

import io
from typing import Callable

from docx import Document
from docx.document import Document as DocumentObject


def assert_round_trip(
    create_fn: Callable[[DocumentObject], None],
    assert_fn: Callable[[DocumentObject], None],
) -> None:
    """Create a document, save to memory, re-open, and run assertions.

    `create_fn` receives a new blank `Document` and should add content to it.
    `assert_fn` receives the re-opened `Document` and should assert correctness.
    """
    # -- create phase --
    doc = Document()
    create_fn(doc)

    # -- save to memory --
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)

    # -- re-open and assert --
    doc2 = Document(stream)
    assert_fn(doc2)


def round_trip_document(doc: DocumentObject) -> DocumentObject:
    """Save a document to memory and re-open it. Returns the re-opened document."""
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return Document(stream)


def save_and_reopen(doc: DocumentObject, path: str) -> DocumentObject:
    """Save a document to a file path and re-open it. Returns the re-opened document.

    Useful for debugging — the saved file can be opened in Word/LibreOffice.
    """
    doc.save(path)
    return Document(path)
