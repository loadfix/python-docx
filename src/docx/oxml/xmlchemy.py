"""Re-export of :mod:`ooxml_xmlchemy`.

The xmlchemy descriptor DSL now lives in the shared
:mod:`ooxml_xmlchemy` package.  This module keeps the existing
``docx.oxml.xmlchemy.*`` import paths working so downstream element-class
modules (and third-party extensions) are not touched.

Namespace-primitive wiring (``qn``, ``nsmap``, ``OxmlElement``,
``clark_to_nsptag``) is performed once in :mod:`docx.oxml.parser` at
import time via :func:`ooxml_xmlchemy.configure_namespace_registry`.
"""

from __future__ import annotations

from ooxml_xmlchemy import (
    AttributeType,
    BaseAttribute,
    BaseOxmlElement,
    Choice,
    MetaOxmlElement,
    OneAndOnlyOne,
    OneOrMore,
    OptionalAttribute,
    RequiredAttribute,
    XmlString,
    ZeroOrMore,
    ZeroOrMoreChoice,
    ZeroOrOne,
    ZeroOrOneChoice,
    lazyproperty,
    serialize_for_reading,
)

# -- internal helpers re-exported for tests and test-only callers --
from ooxml_xmlchemy.xmlchemy import _OxmlElementBase, _XP  # noqa: F401

__all__ = [
    "AttributeType",
    "BaseAttribute",
    "BaseOxmlElement",
    "Choice",
    "MetaOxmlElement",
    "OneAndOnlyOne",
    "OneOrMore",
    "OptionalAttribute",
    "RequiredAttribute",
    "XmlString",
    "ZeroOrMore",
    "ZeroOrMoreChoice",
    "ZeroOrOne",
    "ZeroOrOneChoice",
    "lazyproperty",
    "serialize_for_reading",
]
