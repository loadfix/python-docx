"""Reference file comparison helpers.

Provides utilities for comparing python-docx output against reference .docx files
created in Microsoft Word. This ensures python-docx can correctly read files produced
by Word and that its output is structurally compatible.
"""

from __future__ import annotations

import os

from lxml import etree

from tests.helpers.xmlparse import parse_docx_xml

_REF_DOCS_DIR = os.path.join(os.path.dirname(__file__), "..", "ref-docs")


def ref_docx_path(name: str) -> str:
    """Return the absolute path to a reference .docx file by name (without extension)."""
    return os.path.join(_REF_DOCS_DIR, f"{name}.docx")


def ref_docx_exists(name: str) -> bool:
    """Return True if a reference .docx file with the given name exists."""
    return os.path.exists(ref_docx_path(name))


def compare_xml_structure(
    actual_path: str,
    reference_path: str,
    part_name: str,
    ignore_attrs: set[str] | None = None,
) -> list[str]:
    """Compare the XML structure of a part between two .docx files.

    Returns a list of differences. An empty list means the structures match.
    Only compares element tags and specified attributes — text content and
    element ordering are compared, but whitespace differences are ignored.

    `ignore_attrs` is a set of attribute names (in Clark notation) to exclude
    from comparison. This is useful for attributes like `w:id` that may differ
    between files but are not structurally significant.
    """
    actual_elem = parse_docx_xml(actual_path, part_name)
    ref_elem = parse_docx_xml(reference_path, part_name)

    if actual_elem is None and ref_elem is None:
        return []
    if actual_elem is None:
        return [f"Part '{part_name}' missing in actual file"]
    if ref_elem is None:
        return [f"Part '{part_name}' missing in reference file"]

    ignore = ignore_attrs or set()
    differences: list[str] = []
    _compare_elements(actual_elem, ref_elem, "", ignore, differences)
    return differences


def _compare_elements(
    actual: etree._Element,
    reference: etree._Element,
    path: str,
    ignore_attrs: set[str],
    differences: list[str],
) -> None:
    """Recursively compare two XML elements for structural equivalence."""
    current_path = f"{path}/{_local_tag(actual)}"

    # -- Compare tags --
    if actual.tag != reference.tag:
        differences.append(f"{current_path}: tag mismatch: '{actual.tag}' vs '{reference.tag}'")
        return

    # -- Compare attributes (excluding ignored ones) --
    actual_attrs = {k: v for k, v in actual.attrib.items() if k not in ignore_attrs}
    ref_attrs = {k: v for k, v in reference.attrib.items() if k not in ignore_attrs}
    if actual_attrs != ref_attrs:
        differences.append(
            f"{current_path}: attribute mismatch: {actual_attrs} vs {ref_attrs}"
        )

    # -- Compare text content (stripped) --
    actual_text = (actual.text or "").strip()
    ref_text = (reference.text or "").strip()
    if actual_text != ref_text:
        differences.append(
            f"{current_path}: text mismatch: '{actual_text}' vs '{ref_text}'"
        )

    # -- Compare children --
    actual_children = list(actual)
    ref_children = list(reference)

    if len(actual_children) != len(ref_children):
        differences.append(
            f"{current_path}: child count mismatch: "
            f"{len(actual_children)} vs {len(ref_children)}"
        )
        return

    for a_child, r_child in zip(actual_children, ref_children):
        _compare_elements(a_child, r_child, current_path, ignore_attrs, differences)


def _local_tag(elem: etree._Element) -> str:
    """Return just the local part of an element's tag (strips namespace)."""
    tag = elem.tag
    if isinstance(tag, str) and tag.startswith("{"):
        return tag.split("}", 1)[1]
    return str(tag)
