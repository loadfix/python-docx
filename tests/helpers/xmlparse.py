"""Helpers for extracting and parsing XML from .docx files."""

from __future__ import annotations

import zipfile
from typing import Optional

from lxml import etree


def parse_docx_xml(docx_path: str, part_name: str) -> Optional[etree._Element]:
    """Extract and parse an XML part from a .docx file.

    Returns the parsed lxml Element for the specified part, or None if the part
    does not exist in the archive.

    Args:
        docx_path: Path to the .docx file.
        part_name: The part name within the zip (e.g. "word/comments.xml").
    """
    with zipfile.ZipFile(docx_path) as zf:
        if part_name not in zf.namelist():
            return None
        xml_bytes = zf.read(part_name)
    return etree.fromstring(xml_bytes)


def list_docx_parts(docx_path: str) -> list[str]:
    """Return a list of all part names in a .docx file."""
    with zipfile.ZipFile(docx_path) as zf:
        return zf.namelist()
