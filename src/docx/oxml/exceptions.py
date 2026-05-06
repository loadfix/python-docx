"""Exceptions for oxml sub-package.

:class:`InvalidXmlError` is re-exported from the shared
:mod:`ooxml_xmlchemy.exceptions` module so the ``isinstance`` identity
is stable whether callers import it from docx or from the shared
package.  :class:`XmlchemyError` is kept as an alias of
:class:`ValueError` for callers that had historically caught the parent
class — the shared ``InvalidXmlError`` is itself a subclass of
:class:`ValueError`.
"""

from __future__ import annotations

from ooxml_xmlchemy.exceptions import InvalidXmlError

__all__ = ["InvalidXmlError", "XmlchemyError"]


XmlchemyError = ValueError
