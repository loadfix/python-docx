"""Test helpers for validating python-docx output across multiple layers.

Provides utilities for XML structure validation, OOXML schema validation,
round-trip testing, and reference file comparison.
"""

from tests.helpers.roundtrip import assert_round_trip
from tests.helpers.validate import extract_xml_part, validate_ooxml_structure
from tests.helpers.xmlparse import parse_docx_xml

__all__ = [
    "assert_round_trip",
    "extract_xml_part",
    "parse_docx_xml",
    "validate_ooxml_structure",
]
