"""Re-export of :class:`ooxml_docprops.oxml.CT_CoreProperties`.

Historically ``docx.oxml.coreprops`` defined ``CT_CoreProperties`` inline.
As of 2026.05 the canonical implementation lives in the shared
:mod:`ooxml_docprops.oxml` package; this module keeps the historical
import path working for downstream consumers while preserving two
docx-specific contracts on datetime serialisation:

1. :meth:`_parse_W3CDTF_to_datetime` returns tz-aware UTC
   (``tzinfo=datetime.timezone.utc``); the shared base returns naive UTC.
2. :meth:`_set_element_datetime` converts tz-aware datetimes to UTC
   before writing and writes a trailing ``Z`` (upstream#1542 fix).

.. versionchanged:: 2026.05.0
    Implementation relocated to ``python-ooxml-docprops``; docx-specific
    datetime semantics preserved on the local subclass.
"""

from __future__ import annotations

import datetime as dt

# ---------------------------------------------------------------------------
# Namespace-registry safety: importing ``ooxml_docprops.oxml`` reconfigures
# the process-global ``ooxml_xmlchemy`` namespace registry to the shared
# docprops one. Restore docx's registry before returning so subsequent
# CT_* imports in ``docx.oxml.__init__`` resolve their descriptors against
# the docx registry (which knows ``w:``, ``wp:``, ``m:``, ... prefixes).
# ---------------------------------------------------------------------------
from ooxml_docprops.oxml import (
    CT_CoreProperties as _CT_CoreProperties_Base,
    qn as _shared_qn,
)
from ooxml_xmlchemy import configure_namespace_registry as _configure

from docx.oxml.parser import _DocxNamespaceRegistry as _DocxRegistry


class CT_CoreProperties(_CT_CoreProperties_Base):
    """docx flavour of :class:`ooxml_docprops.oxml.CT_CoreProperties`.

    Preserves two docx-specific datetime contracts on top of the shared
    base (see module docstring for the rationale).
    """

    @classmethod
    def new(cls) -> "CT_CoreProperties":
        """Return a new empty ``<cp:coreProperties>`` element (docx flavour).

        Routes parsing through ``docx.oxml.parse_xml`` so the returned
        element is an instance of *this* (docx) subclass rather than the
        shared base. The shared base's ``new()`` is a ``@staticmethod`` and
        uses the shared parser, which would return a plain-base instance
        without docx's datetime overrides.
        """
        from docx.oxml.parser import parse_xml as _docx_parse_xml
        from typing import cast

        element = cast(
            "CT_CoreProperties",
            _docx_parse_xml(cls._coreProperties_tmpl),
        )
        return element

    @classmethod
    def _parse_W3CDTF_to_datetime(cls, w3cdtf_str: str) -> dt.datetime:
        """Parse W3CDTF text and return a ``tz-aware UTC`` :class:`datetime`.

        The shared base returns the same instant with ``tzinfo=None``
        (naive UTC). Tag the return value with ``timezone.utc`` so callers
        that compare against aware constants (the docx test suite does)
        continue to match.
        """
        value = super()._parse_W3CDTF_to_datetime(w3cdtf_str)
        if value.tzinfo is None:
            value = value.replace(tzinfo=dt.timezone.utc)
        return value

    def _set_element_datetime(self, prop_name: str, value: dt.datetime) -> None:
        """Serialise *value* as W3CDTF, converting tz-aware values to UTC.

        Regression for upstream#1542: naive values are assumed to already
        be UTC and serialised directly; aware values are normalised to UTC
        before the trailing ``Z`` is written, so the on-disk instant is
        honest (rather than mislabelling local wall-clock time as UTC).
        """
        if not isinstance(value, dt.datetime):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise ValueError(
                "property requires <type 'datetime.datetime'> object, got %s"
                % type(value)
            )
        if value.tzinfo is not None:
            value = value.astimezone(dt.timezone.utc)
        element = self._get_or_add(prop_name)
        element.text = value.strftime("%Y-%m-%dT%H:%M:%SZ")
        if prop_name in ("created", "modified"):
            element.set(_shared_qn("xsi:type"), "dcterms:W3CDTF")


_configure(_DocxRegistry())

__all__ = ["CT_CoreProperties"]
