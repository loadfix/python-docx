"""pytest fixtures that are shared across test modules."""

from __future__ import annotations

import io
import os
import tempfile
import zipfile
from typing import TYPE_CHECKING

import pytest

from docx import Document
from docx.document import Document as DocumentCls
from docx.opc.strict import STRICT_TO_TRANSITIONAL

if TYPE_CHECKING:
    from docx import types as t
    from docx.parts.story import StoryPart


def _rewrite_ns_to_strict(pkg_bytes: bytes) -> bytes:
    """Rewrite every Transitional OOXML namespace URI in `pkg_bytes` to its
    ECMA-376 Strict counterpart.

    Operates at the ZIP level — each XML part (and ``[Content_Types].xml`` /
    ``*.rels``) has its bytes scanned and substituted in place; binary parts
    (images, etc.) are passed through unchanged.

    Used by the R16-1 strict-fixture round-trip tests to synthesize a real
    Strict ``.docx`` from python-docx's minimal Transitional template — this
    lets the tests exercise the open-time namespace sniff, the strict flag
    plumbing, and an edit + save + reopen cycle end-to-end without needing
    a hand-crafted Strict fixture file in the repo.
    """
    trans_to_strict = {t: s for s, t in STRICT_TO_TRANSITIONAL.items()}
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(pkg_bytes)) as zin:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if (
                    info.filename.endswith((".xml", ".rels"))
                    or info.filename == "[Content_Types].xml"
                ):
                    for trans, strict in trans_to_strict.items():
                        data = data.replace(trans, strict)
                zout.writestr(info, data)
    return out.getvalue()


@pytest.fixture
def fake_parent() -> t.ProvidesStoryPart:
    class ProvidesStoryPart:
        @property
        def part(self) -> StoryPart:
            raise NotImplementedError

    return ProvidesStoryPart()


@pytest.fixture
def tmp_docx_path():
    """Yield a temporary file path for .docx output; cleaned up after test."""
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    yield path
    if os.path.exists(path):
        os.unlink(path)


@pytest.fixture
def blank_document() -> DocumentCls:
    """Return a new blank Document for use in tests."""
    return Document()
