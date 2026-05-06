"""Re-export of :mod:`ooxml_opc.oxml` with docx-local helpers.

The :class:`CT_Types` / :class:`CT_Relationships` / :class:`CT_Default` /
:class:`CT_Override` / :class:`CT_Relationship` xmlchemy element classes
live in :mod:`ooxml_opc.oxml` and are re-exported here so external
callers (including test code) continue to import them from this module.

docx-local helpers retained:

* :func:`qn` — restricted to the OPC namespace family.
* :func:`serialize_for_reading` — pretty-printed str for tests (kept in
  addition to the shared :func:`ooxml_opc.oxml.serialize_for_reading`).
* :data:`nsmap` — OPC-local prefix → URI mapping.
* :data:`oxml_parser` / :data:`element_class_lookup` / legacy parser
  handles — re-exported from the shared runtime for backward-compat.
"""

from __future__ import annotations

from lxml import etree

from ooxml_opc.oxml import (  # noqa: F401 -- re-exports
    BaseOxmlElement,
    CT_Default,
    CT_Override,
    CT_Relationship,
    CT_Relationships,
    CT_Types,
    parse_xml,
    serialize_for_reading,
    serialize_part_xml,
)
from ooxml_opc.constants import NAMESPACE as NS

__all__ = [
    "BaseOxmlElement",
    "CT_Default",
    "CT_Override",
    "CT_Relationship",
    "CT_Relationships",
    "CT_Types",
    "nsmap",
    "parse_xml",
    "qn",
    "serialize_for_reading",
    "serialize_part_xml",
]


#: OPC-local nsmap. Kept module-level so callers that built it into their
#: own serialisation paths continue to see the same structure.
nsmap = {
    "ct": NS.OPC_CONTENT_TYPES,
    "pr": NS.OPC_RELATIONSHIPS,
    "r": NS.OFC_RELATIONSHIPS,
}


def qn(tag: str) -> str:
    """Turn a namespace-prefixed tag name into Clark-notation.

    For example, ``qn('pr:Relationship')`` returns
    ``'{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'``.
    Restricted to the OPC namespace family (``ct:`` / ``pr:`` / ``r:``) — the
    general-purpose docx xmlchemy nsmap lives in :mod:`docx.oxml.ns`.
    """
    prefix, tagroot = tag.split(":")
    uri = nsmap[prefix]
    return "{%s}%s" % (uri, tagroot)


# ---------------------------------------------------------------------------
# docx-local ``.new()`` factory methods patched onto the shared classes
# ---------------------------------------------------------------------------
# docx callers historically constructed CT_Default / CT_Override via static
# ``.new(...)`` factories. The shared runtime exposes these constructions via
# ``CT_Types.add_default()`` / ``.add_override()`` instead. Patch the legacy
# factories back on for backward-compatibility.


def _ct_default_new(ext: str, content_type: str):
    xml = '<Default xmlns="%s"/>' % nsmap["ct"]
    default = parse_xml(xml)
    default.set("Extension", ext)
    default.set("ContentType", content_type)
    return default


def _ct_override_new(partname: str, content_type: str):
    xml = '<Override xmlns="%s"/>' % nsmap["ct"]
    override = parse_xml(xml)
    override.set("PartName", partname)
    override.set("ContentType", content_type)
    return override


# -- Only patch if not already present to play nicely with any parent
# -- library that registers an override. --
if not hasattr(CT_Default, "new") or not callable(getattr(CT_Default, "new", None)):
    CT_Default.new = staticmethod(_ct_default_new)  # type: ignore[attr-defined]
if not hasattr(CT_Override, "new") or not callable(getattr(CT_Override, "new", None)):
    CT_Override.new = staticmethod(_ct_override_new)  # type: ignore[attr-defined]
