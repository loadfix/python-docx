"""Exceptions for oxml sub-package."""
from __future__ import annotations


class XmlchemyError(Exception):
    """Generic error class."""


class InvalidXmlError(XmlchemyError):
    """Raised when invalid XML is encountered, such as on attempt to access a missing
    required child element."""
