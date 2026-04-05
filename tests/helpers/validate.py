"""OOXML structure and schema validation helpers for .docx files."""

from __future__ import annotations

import os
import zipfile
from typing import Sequence

from lxml import etree

from tests.helpers.xmlparse import parse_docx_xml

# -- Namespaces used in OOXML documents ------------------------------------------------

_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class OoxmlValidationError(Exception):
    """Raised when OOXML structural validation fails."""


def extract_xml_part(docx_path: str, part_name: str) -> etree._Element:
    """Extract and parse an XML part from a .docx, raising if it does not exist.

    This is a convenience wrapper around `parse_docx_xml` that raises rather than
    returning None when the part is missing.
    """
    element = parse_docx_xml(docx_path, part_name)
    if element is None:
        raise OoxmlValidationError(f"Part '{part_name}' not found in {docx_path}")
    return element


def validate_ooxml_structure(docx_path: str) -> list[str]:
    """Validate the structural integrity of a .docx file.

    Returns a list of validation error messages. An empty list means the file is
    structurally valid. Checks include:

    - The file is a valid ZIP archive.
    - `[Content_Types].xml` exists and is well-formed XML.
    - Every Override in `[Content_Types].xml` references a part that exists.
    - `_rels/.rels` exists and is well-formed XML.
    - `word/document.xml` exists and has a `w:document` root element.
    - All relationship targets in `word/_rels/document.xml.rels` exist in the archive.
    - All XML parts referenced are well-formed XML.
    """
    errors: list[str] = []

    # -- Check that it's a valid zip -------------------------------------------------
    if not zipfile.is_zipfile(docx_path):
        return [f"{docx_path} is not a valid ZIP file"]

    with zipfile.ZipFile(docx_path) as zf:
        names = set(zf.namelist())

        # -- [Content_Types].xml -----------------------------------------------------
        if "[Content_Types].xml" not in names:
            errors.append("Missing [Content_Types].xml")
        else:
            ct_elem = _parse_zip_xml(zf, "[Content_Types].xml", errors)
            if ct_elem is not None:
                _check_content_types_overrides(ct_elem, names, errors)

        # -- _rels/.rels -------------------------------------------------------------
        if "_rels/.rels" not in names:
            errors.append("Missing _rels/.rels")
        else:
            _parse_zip_xml(zf, "_rels/.rels", errors)

        # -- word/document.xml -------------------------------------------------------
        if "word/document.xml" not in names:
            errors.append("Missing word/document.xml")
        else:
            doc_elem = _parse_zip_xml(zf, "word/document.xml", errors)
            if doc_elem is not None:
                _check_root_tag(doc_elem, f"{{{_WML_NS}}}document", "word/document.xml", errors)

        # -- word/_rels/document.xml.rels --------------------------------------------
        doc_rels_path = "word/_rels/document.xml.rels"
        if doc_rels_path in names:
            rels_elem = _parse_zip_xml(zf, doc_rels_path, errors)
            if rels_elem is not None:
                _check_relationship_targets(rels_elem, names, errors)

        # -- Validate all XML parts are well-formed ----------------------------------
        already_parsed = {"[Content_Types].xml", "_rels/.rels", "word/document.xml", doc_rels_path}
        for name in names:
            if name in already_parsed:
                continue
            if name.endswith(".xml") or name.endswith(".rels"):
                _parse_zip_xml(zf, name, errors)

    return errors


def validate_content_type_present(docx_path: str, content_type: str) -> bool:
    """Return True if `content_type` is registered in [Content_Types].xml."""
    ct_elem = extract_xml_part(docx_path, "[Content_Types].xml")
    for override in ct_elem.findall(f"{{{_CONTENT_TYPES_NS}}}Override"):
        if override.get("ContentType") == content_type:
            return True
    for default in ct_elem.findall(f"{{{_CONTENT_TYPES_NS}}}Default"):
        if default.get("ContentType") == content_type:
            return True
    return False


def validate_relationship_present(
    docx_path: str,
    rel_type: str,
    rels_part: str = "word/_rels/document.xml.rels",
) -> bool:
    """Return True if a relationship of `rel_type` exists in the specified rels part."""
    rels_elem = parse_docx_xml(docx_path, rels_part)
    if rels_elem is None:
        return False
    for rel in rels_elem.findall(f"{{{_RELS_NS}}}Relationship"):
        if rel.get("Type") == rel_type:
            return True
    return False


def validate_elements_present(
    docx_path: str,
    part_name: str,
    xpath: str,
    namespaces: dict[str, str] | None = None,
    min_count: int = 1,
) -> list[etree._Element]:
    """Assert that at least `min_count` elements matching `xpath` exist in `part_name`.

    Returns the matching elements. Raises OoxmlValidationError if the count is below
    `min_count`.
    """
    element = extract_xml_part(docx_path, part_name)
    ns = namespaces or {"w": _WML_NS}
    matches = element.xpath(xpath, namespaces=ns)
    if not isinstance(matches, list):
        matches = [matches]
    if len(matches) < min_count:
        raise OoxmlValidationError(
            f"Expected at least {min_count} elements matching '{xpath}' in "
            f"'{part_name}', found {len(matches)}"
        )
    return matches


# -- internal helpers ----------------------------------------------------------------


def _parse_zip_xml(
    zf: zipfile.ZipFile, name: str, errors: list[str]
) -> etree._Element | None:
    """Parse an XML file from the zip, appending to errors on failure."""
    try:
        return etree.fromstring(zf.read(name))
    except etree.XMLSyntaxError as e:
        errors.append(f"Malformed XML in {name}: {e}")
        return None


def _check_content_types_overrides(
    ct_elem: etree._Element, archive_names: set[str], errors: list[str]
) -> None:
    """Verify every Override PartName in [Content_Types].xml has a matching archive entry."""
    for override in ct_elem.findall(f"{{{_CONTENT_TYPES_NS}}}Override"):
        part_name = override.get("PartName", "")
        # PartName starts with "/" in the XML, but zip entries don't
        zip_name = part_name.lstrip("/")
        if zip_name not in archive_names:
            errors.append(
                f"[Content_Types].xml Override references missing part: {part_name}"
            )


def _check_root_tag(
    elem: etree._Element, expected_tag: str, part_name: str, errors: list[str]
) -> None:
    """Verify an element has the expected root tag."""
    if elem.tag != expected_tag:
        errors.append(
            f"{part_name}: expected root tag '{expected_tag}', got '{elem.tag}'"
        )


def _check_relationship_targets(
    rels_elem: etree._Element, archive_names: set[str], errors: list[str]
) -> None:
    """Verify relationship targets exist in the archive (for internal targets only)."""
    for rel in rels_elem.findall(f"{{{_RELS_NS}}}Relationship"):
        target_mode = rel.get("TargetMode", "Internal")
        if target_mode == "External":
            continue
        target = rel.get("Target", "")
        # Relationship targets are relative to the source part's directory
        if target.startswith("/"):
            zip_path = target.lstrip("/")
        else:
            zip_path = f"word/{target}"
        # Normalize parent-directory references (e.g. "word/../customXml/item1.xml")
        zip_path = os.path.normpath(zip_path).replace("\\", "/")
        if zip_path not in archive_names:
            errors.append(
                f"Relationship target '{target}' not found in archive (expected '{zip_path}')"
            )
