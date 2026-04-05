"""OOXML schema validation using lxml.etree.XMLSchema.

Validates individual XML parts against XSD schemas derived from ECMA-376.
The schemas are simplified subsets focusing on the elements python-docx produces.

For full schema validation, the complete ECMA-376 XSD files can be downloaded from:
https://www.ecma-international.org/publications-and-standards/standards/ecma-376/

This module provides a practical alternative that validates the most important
structural constraints without requiring the full (very large) schema set.
"""

from __future__ import annotations

import os
import zipfile
from typing import Optional

from lxml import etree

_SCHEMAS_DIR = os.path.join(os.path.dirname(__file__), "schemas")

# -- OOXML namespace URIs --
WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

# -- Namespace map for xpath queries --
OOXML_NSMAP = {
    "w": WML_NS,
    "r": REL_NS,
    "pr": PKG_REL_NS,
    "ct": CT_NS,
}


class SchemaValidationResult:
    """Result of validating an XML part against a schema."""

    def __init__(self, is_valid: bool, errors: list[str]):
        self.is_valid = is_valid
        self.errors = errors

    def __bool__(self) -> bool:
        return self.is_valid

    def __repr__(self) -> str:
        if self.is_valid:
            return "SchemaValidationResult(valid)"
        return f"SchemaValidationResult(invalid, {len(self.errors)} errors)"


def validate_part_xml(
    xml_bytes: bytes,
    schema: etree.XMLSchema,
) -> SchemaValidationResult:
    """Validate XML bytes against the provided lxml XMLSchema.

    Returns a SchemaValidationResult with is_valid=True if the XML is valid,
    or is_valid=False with a list of error messages otherwise.
    """
    try:
        doc = etree.fromstring(xml_bytes)
    except etree.XMLSyntaxError as e:
        return SchemaValidationResult(False, [f"XML syntax error: {e}"])

    is_valid = schema.validate(doc)
    errors = [str(e) for e in schema.error_log] if not is_valid else []
    return SchemaValidationResult(is_valid, errors)


def load_schema(schema_path: str) -> etree.XMLSchema:
    """Load an XSD schema from a file path."""
    with open(schema_path, "rb") as f:
        schema_doc = etree.parse(f)
    return etree.XMLSchema(schema_doc)


def load_bundled_schema(name: str) -> Optional[etree.XMLSchema]:
    """Load a bundled XSD schema by name.

    Returns None if the schema file does not exist (schemas are optional and may
    need to be downloaded separately).
    """
    path = os.path.join(_SCHEMAS_DIR, f"{name}.xsd")
    if not os.path.exists(path):
        return None
    return load_schema(path)


def validate_docx_xml_parts(docx_path: str) -> dict[str, SchemaValidationResult]:
    """Validate all XML parts in a .docx file for well-formedness.

    This is a lighter check that ensures every XML part in the archive is at least
    well-formed XML. For schema validation of specific parts, use `validate_part_xml`
    with an appropriate schema.

    Returns a dict mapping part names to their validation results.
    """
    results: dict[str, SchemaValidationResult] = {}

    with zipfile.ZipFile(docx_path) as zf:
        for name in zf.namelist():
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            xml_bytes = zf.read(name)
            try:
                etree.fromstring(xml_bytes)
                results[name] = SchemaValidationResult(True, [])
            except etree.XMLSyntaxError as e:
                results[name] = SchemaValidationResult(False, [f"XML syntax error: {e}"])

    return results
