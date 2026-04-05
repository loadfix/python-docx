"""Layer 2: OOXML schema validation helpers.

Validates XML parts from .docx files against the ECMA-376 OOXML XSD schemas.

Schema files are expected in `tests/schema/` directory. Use the download script
`tests/helpers/download_schema.py` to fetch them from the ECMA-376 standard.
"""

from __future__ import annotations

import os
import zipfile
from typing import Sequence

from lxml import etree

_SCHEMA_DIR = os.path.join(os.path.dirname(__file__), "..", "schema")

# -- mapping from .docx part paths to their primary XSD schema file --
_PART_SCHEMA_MAP: dict[str, str] = {
    "word/document.xml": "wml.xsd",
    "word/comments.xml": "wml.xsd",
    "word/footnotes.xml": "wml.xsd",
    "word/endnotes.xml": "wml.xsd",
    "word/numbering.xml": "wml.xsd",
    "word/styles.xml": "wml.xsd",
    "word/settings.xml": "wml.xsd",
}


def _schema_path(schema_name: str) -> str:
    return os.path.join(_SCHEMA_DIR, schema_name)


def _schema_available() -> bool:
    """Return True if at least the primary WML schema file is available."""
    return os.path.isfile(_schema_path("wml.xsd"))


def load_schema(schema_name: str = "wml.xsd") -> etree.XMLSchema | None:
    """Load and return an XMLSchema from the schema directory.

    Returns None if the schema file is not available. This allows tests
    to be skipped gracefully when schemas haven't been downloaded.
    """
    path = _schema_path(schema_name)
    if not os.path.isfile(path):
        return None
    schema_doc = etree.parse(path)
    return etree.XMLSchema(schema_doc)


def validate_part_xml(
    docx_path: str,
    part_name: str,
    schema: etree.XMLSchema | None = None,
) -> list[str]:
    """Validate a single XML part from a .docx against an OOXML schema.

    Returns a list of validation error messages (empty if valid).
    If schema is None and no default schema is available, returns a single
    message indicating schemas are not installed.
    """
    if schema is None:
        schema_name = _PART_SCHEMA_MAP.get(part_name, "wml.xsd")
        schema = load_schema(schema_name)

    if schema is None:
        return [f"Schema not available for '{part_name}' — run download_schema.py"]

    with zipfile.ZipFile(docx_path, "r") as zf:
        try:
            xml_bytes = zf.read(part_name)
        except KeyError:
            return [f"Part '{part_name}' not found in archive"]

    doc = etree.fromstring(xml_bytes)
    if schema.validate(doc):
        return []

    return [str(err) for err in schema.error_log]


def validate_ooxml(
    docx_path: str,
    parts: Sequence[str] | None = None,
) -> list[str]:
    """Validate one or more XML parts from a .docx file against OOXML schemas.

    If `parts` is None, validates all known WML parts found in the archive.
    Returns a list of validation error messages (empty if all valid).
    """
    if not _schema_available():
        return ["OOXML schemas not available — run download_schema.py to fetch them"]

    with zipfile.ZipFile(docx_path, "r") as zf:
        archive_entries = set(zf.namelist())

    if parts is None:
        parts = [p for p in _PART_SCHEMA_MAP if p in archive_entries]

    errors: list[str] = []
    for part_name in parts:
        part_errors = validate_part_xml(docx_path, part_name)
        errors.extend(part_errors)

    return errors
