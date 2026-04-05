"""Exceptions specific to python-opc.

The base exception class is OpcError.
"""
from __future__ import annotations


class OpcError(Exception):
    """Base error class for python-opc."""


class PackageNotFoundError(OpcError):
    """Raised when a package cannot be found at the specified path."""
