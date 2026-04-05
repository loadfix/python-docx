"""Layer 4: Reference file comparison helpers.

Provides utilities for reading and comparing against reference .docx files
created in Microsoft Word.
"""

from __future__ import annotations

import os
import zipfile

from lxml import etree

_REF_DOCS_DIR = os.path.join(os.path.dirname(__file__), "..", "ref-docs")


def ref_docx_path(name: str) -> str:
    """Return the absolute path to a reference .docx file by stem name.

    For example, `ref_docx_path("comments-simple")` returns the path to
    `tests/ref-docs/comments-simple.docx`.
    """
    return os.path.join(_REF_DOCS_DIR, f"{name}.docx")


def ref_docx_exists(name: str) -> bool:
    """Return True if the named reference .docx file exists."""
    return os.path.isfile(ref_docx_path(name))


def extract_ref_xml(name: str, part_name: str) -> etree._Element:
    """Extract and parse an XML part from a reference .docx file.

    `name` is the stem name of the .docx file in `tests/ref-docs/`.
    `part_name` is the zip entry path (e.g. "word/comments.xml").
    """
    path = ref_docx_path(name)
    with zipfile.ZipFile(path, "r") as zf:
        xml_bytes = zf.read(part_name)
    return etree.fromstring(xml_bytes)


def compare_xml_structure(
    actual: etree._Element,
    expected: etree._Element,
    ignore_attrs: set[str] | None = None,
) -> list[str]:
    """Compare two XML element trees structurally.

    Checks tag names, attributes, text content, and child structure.
    Returns a list of difference descriptions (empty if structurally equivalent).

    `ignore_attrs` is an optional set of attribute names to skip during comparison
    (useful for ignoring auto-generated IDs or timestamps).
    """
    differences: list[str] = []
    _compare_elements(actual, expected, "", differences, ignore_attrs or set())
    return differences


def _compare_elements(
    actual: etree._Element,
    expected: etree._Element,
    path: str,
    differences: list[str],
    ignore_attrs: set[str],
) -> None:
    """Recursively compare two elements and collect differences."""
    current_path = f"{path}/{actual.tag}"

    # -- compare tags --
    if actual.tag != expected.tag:
        differences.append(f"Tag mismatch at {path}: {actual.tag} != {expected.tag}")
        return

    # -- compare attributes --
    actual_attrs = {k: v for k, v in actual.attrib.items() if k not in ignore_attrs}
    expected_attrs = {k: v for k, v in expected.attrib.items() if k not in ignore_attrs}
    if actual_attrs != expected_attrs:
        differences.append(
            f"Attribute mismatch at {current_path}: {actual_attrs} != {expected_attrs}"
        )

    # -- compare text --
    actual_text = (actual.text or "").strip()
    expected_text = (expected.text or "").strip()
    if actual_text != expected_text:
        differences.append(
            f"Text mismatch at {current_path}: '{actual_text}' != '{expected_text}'"
        )

    # -- compare children count --
    actual_children = list(actual)
    expected_children = list(expected)
    if len(actual_children) != len(expected_children):
        differences.append(
            f"Child count mismatch at {current_path}: "
            f"{len(actual_children)} != {len(expected_children)}"
        )
        return

    # -- recurse into children --
    for act_child, exp_child in zip(actual_children, expected_children):
        _compare_elements(act_child, exp_child, current_path, differences, ignore_attrs)
