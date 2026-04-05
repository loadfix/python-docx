"""Layer 1: XML structure validation helpers.

Extracts and validates XML parts from .docx files for structural correctness.
"""

from __future__ import annotations

import zipfile
from typing import Any

from lxml import etree

# -- OOXML namespace map for XPath queries --
OOXML_NSMAP: dict[str, str] = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}


def extract_xml_part(docx_path: str, part_name: str) -> etree._Element:
    """Extract and parse an XML part from a .docx file.

    `part_name` is the path within the zip archive, e.g. "word/comments.xml",
    "word/document.xml", "[Content_Types].xml", or "word/_rels/document.xml.rels".

    Returns the parsed lxml Element tree root.

    Raises `KeyError` if the part does not exist in the archive.
    """
    with zipfile.ZipFile(docx_path, "r") as zf:
        xml_bytes = zf.read(part_name)
    return etree.fromstring(xml_bytes)


def list_zip_entries(docx_path: str) -> list[str]:
    """Return the list of entry names in the .docx zip archive."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        return zf.namelist()


def validate_xml_structure(docx_path: str, checks: list[dict[str, Any]]) -> list[str]:
    """Run a list of structural checks against a .docx file.

    Each check is a dict with:
      - "part": the XML part path (e.g. "word/comments.xml")
      - "xpath": an XPath expression to evaluate
      - "expected_count": (optional) expected number of matching nodes
      - "expected_min": (optional) minimum number of matching nodes
      - "expected_values": (optional) list of expected text/attribute values

    Returns a list of failure messages. An empty list means all checks passed.
    """
    failures: list[str] = []
    for check in checks:
        part_name = check["part"]
        xpath_expr = check["xpath"]

        try:
            root = extract_xml_part(docx_path, part_name)
        except KeyError:
            failures.append(f"Part '{part_name}' not found in archive")
            continue

        results = root.xpath(xpath_expr, namespaces=OOXML_NSMAP)

        if "expected_count" in check:
            expected = check["expected_count"]
            if len(results) != expected:
                failures.append(
                    f"XPath '{xpath_expr}' in '{part_name}': "
                    f"expected {expected} matches, got {len(results)}"
                )

        if "expected_min" in check:
            expected_min = check["expected_min"]
            if len(results) < expected_min:
                failures.append(
                    f"XPath '{xpath_expr}' in '{part_name}': "
                    f"expected at least {expected_min} matches, got {len(results)}"
                )

        if "expected_values" in check:
            actual_values = [
                r.text if isinstance(r, etree._Element) else str(r) for r in results
            ]
            expected_values = check["expected_values"]
            if actual_values != expected_values:
                failures.append(
                    f"XPath '{xpath_expr}' in '{part_name}': "
                    f"expected values {expected_values}, got {actual_values}"
                )

    return failures


def assert_part_exists(docx_path: str, part_name: str) -> None:
    """Assert that a named part exists in the .docx zip archive."""
    entries = list_zip_entries(docx_path)
    assert part_name in entries, (
        f"Expected part '{part_name}' not found in archive. "
        f"Available parts: {sorted(entries)}"
    )


def assert_content_type_exists(docx_path: str, content_type: str) -> None:
    """Assert that [Content_Types].xml declares the given content type."""
    root = extract_xml_part(docx_path, "[Content_Types].xml")
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    overrides = root.findall(f"{{{ct_ns}}}Override")
    found_types = [o.get("ContentType") for o in overrides]
    assert content_type in found_types, (
        f"Content type '{content_type}' not found in [Content_Types].xml. "
        f"Found: {found_types}"
    )


def assert_relationship_exists(
    docx_path: str,
    rels_part: str,
    relationship_type: str,
) -> None:
    """Assert that a relationship of the given type exists in the specified .rels part."""
    root = extract_xml_part(docx_path, rels_part)
    rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rels = root.findall(f"{{{rels_ns}}}Relationship")
    found_types = [r.get("Type") for r in rels]
    assert relationship_type in found_types, (
        f"Relationship type '{relationship_type}' not found in '{rels_part}'. "
        f"Found: {found_types}"
    )
