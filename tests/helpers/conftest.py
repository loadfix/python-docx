"""Shared pytest fixtures for the multi-layered testing strategy."""

from __future__ import annotations

import os
import shutil
import tempfile

import pytest

from docx import Document
from docx.document import Document as DocumentCls


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


@pytest.fixture
def libreoffice_available() -> bool:
    """Return True if LibreOffice is available on the system."""
    return shutil.which("libreoffice") is not None
