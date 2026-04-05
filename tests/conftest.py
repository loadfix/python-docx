"""pytest fixtures that are shared across test modules."""

from __future__ import annotations

import os
import tempfile
from typing import TYPE_CHECKING

import pytest

from docx import Document
from docx.document import Document as DocumentCls

if TYPE_CHECKING:
    from docx import types as t
    from docx.parts.story import StoryPart


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
