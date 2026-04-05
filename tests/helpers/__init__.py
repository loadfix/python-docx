"""Test helpers for validating python-docx output across multiple testing layers.

Provides utilities for:
- XML structure validation (Layer 1)
- OOXML schema validation (Layer 2)
- Round-trip testing (Layer 3)
- Reference file comparison (Layer 4)
- LibreOffice headless validation (Layer 5)
"""

from tests.helpers.roundtrip import assert_round_trip
from tests.helpers.xml_structure import extract_xml_part, validate_xml_structure

__all__ = [
    "assert_round_trip",
    "extract_xml_part",
    "validate_xml_structure",
]
