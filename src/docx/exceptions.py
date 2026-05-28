"""Exceptions used with python-docx.

The base exception class is :class:`PythonDocxError`. Structured,
LLM-friendly errors derive from :class:`DocxError` and carry a
machine-readable error ``code``, a human ``message``, an optional
``suggestion`` (e.g. a fuzzy-matched "did you mean" hint), an
optional ``location`` (the conceptual address that triggered the
failure — XPath, paragraph index, dictionary key…) and the
``operation`` that raised. Subclasses also inherit from the legacy
built-in (``KeyError`` / ``ValueError`` / ``IndexError`` /
``RuntimeError``) so existing ``except`` blocks continue to catch
them after the upgrade.

Error catalog
=============

The catalog below lists every machine-readable ``code`` the library
emits. Codes are stable across releases — once exposed, a code's
meaning never changes (we add new codes for new failure modes).

============================  ==============================================
code                          meaning
============================  ==============================================
``STYLE_NOT_FOUND``           A style lookup by name (or
                              ``WD_BUILTIN_STYLE``) failed. The
                              ``suggestion`` attribute carries the closest
                              matching name(s) found in the document.
``STYLE_DUPLICATE``           Attempted to add a style whose name already
                              exists. Suggestion: pick a unique name or
                              update the existing style.
``STYLE_TYPE_MISMATCH``       The style is the wrong ``WD_STYLE_TYPE`` for
                              the operation (e.g. assigning a paragraph
                              style where a character style was expected).
``LATENT_STYLE_NOT_FOUND``    A latent style with the requested name was
                              not found.
``BUILTIN_STYLE_NOT_FOUND``   ``Styles.import_builtin`` could not find the
                              requested built-in style in the bundled
                              templates.
``BOOKMARK_NOT_FOUND``        Lookup or removal of a bookmark by name
                              failed. Suggestion lists nearby names.
``FONT_NOT_FOUND``            ``FontTable[name]`` lookup failed.
``FONT_FAMILY_INVALID``       ``FontTable.add_embedded_font`` rejected an
                              unknown font-family token.
``FONT_EMBED_EMPTY``          ``FontTable.embed_font`` was called with no
                              variant bytes — at least one of regular /
                              bold / italic / bold_italic is required.
``THEME_TOKEN_INVALID``       ``ThemeColors[token]`` was given a token that
                              is not one of the twelve OOXML scheme slots.
``INVALID_COLOR``             RGB color components must be integers in
                              0-255, or a 3- or 6-character hex string.
``INVALID_BRIGHTNESS``        Theme-color brightness must be in
                              ``-1.0 .. +1.0``.
``BRIGHTNESS_NO_THEME``       Cannot set brightness when no theme color is
                              assigned to the run.
``COLUMN_WIDTHS_LENGTH``      ``Section.set_columns`` was given a
                              ``widths`` list whose length disagrees with
                              ``count``.
``BORDER_SIDE_INVALID``       ``Section.set_page_border`` requires
                              ``side`` to be one of ``"top"``,
                              ``"bottom"``, ``"left"``, ``"right"``.
``WATERMARK_LAYOUT_INVALID``  ``Section.add_text_watermark`` requires
                              ``layout`` to be ``"diagonal"`` or
                              ``"horizontal"``.
``SECTION_INDEX``             ``Sections.pop`` index is out of range.
``NOT_A_WORD_FILE``           The file's content type is not a recognised
                              WordprocessingML document or template.
``NOT_A_WORD_TEMPLATE``       The package opened by ``from_template`` is
                              not a ``.dotx`` / ``.dotm`` template.
============================  ==============================================

This catalog is the authoritative list. New codes added in later
versions will appear at the bottom of the file's
``__all_codes__`` tuple.
"""

from __future__ import annotations

import difflib
from typing import Iterable, Optional, Tuple, Type, TypeVar, Union


__all__ = [
    "BookmarkNotFoundError",
    "BuiltinStyleNotFoundError",
    "DocxError",
    "EncryptedDocumentError",
    "FontEmbedEmptyError",
    "FontFamilyInvalidError",
    "FontNotFoundError",
    "InvalidBrightnessError",
    "InvalidColorError",
    "InvalidSpanError",
    "InvalidXmlError",
    "LatentStyleNotFoundError",
    "NestedSectionError",
    "NotAWordFileError",
    "NotAWordTemplateError",
    "OutOfRangeError",
    "PythonDocxError",
    "RmsProtectedDocumentError",
    "StyleDuplicateError",
    "StyleNotFoundError",
    "StyleTypeMismatchError",
    "ThemeTokenInvalidError",
    "ValueOutOfRangeError",
    "closest_names",
]


# -- ``__all_codes__`` is the authoritative list of stable error codes; --
# -- adding a code here is the public-API gesture that "this code is now --
# -- documented and stable". --
__all_codes__: Tuple[str, ...] = (
    "STYLE_NOT_FOUND",
    "STYLE_DUPLICATE",
    "STYLE_TYPE_MISMATCH",
    "LATENT_STYLE_NOT_FOUND",
    "BUILTIN_STYLE_NOT_FOUND",
    "BOOKMARK_NOT_FOUND",
    "FONT_NOT_FOUND",
    "FONT_FAMILY_INVALID",
    "FONT_EMBED_EMPTY",
    "THEME_TOKEN_INVALID",
    "INVALID_COLOR",
    "INVALID_BRIGHTNESS",
    "BRIGHTNESS_NO_THEME",
    "COLUMN_WIDTHS_LENGTH",
    "BORDER_SIDE_INVALID",
    "WATERMARK_LAYOUT_INVALID",
    "SECTION_INDEX",
    "NOT_A_WORD_FILE",
    "NOT_A_WORD_TEMPLATE",
)


class PythonDocxError(Exception):
    """Generic error class — base of every exception raised by python-docx.

    Predates the structured :class:`DocxError` taxonomy; retained as the
    common ancestor so callers can still write ``except PythonDocxError``
    to catch any library-emitted error regardless of its built-in
    co-base.
    """


class DocxError(PythonDocxError):
    """Structured, LLM-friendly base error.

    Each instance carries a stable ``code`` (machine-readable), a human
    ``message``, an optional ``suggestion`` (e.g. a fuzzy "did you
    mean…" hint), an optional ``location`` (the conceptual address
    that triggered the error — XPath, paragraph index, dict key, …)
    and the ``operation`` that raised (typically
    ``"Class.method"`` or ``"module.function"``).

    Subclasses should also inherit from the legacy built-in exception
    type the operation used to raise (for example,
    :class:`StyleNotFoundError` extends both :class:`DocxError` and
    :class:`KeyError` so existing ``except KeyError:`` callers still
    work after the upgrade).

    .. versionadded:: 2026.05.13
    """

    #: The default error ``code`` used when a subclass does not set one.
    #: ``DocxError`` itself is rarely raised directly; subclasses should
    #: override this with a value drawn from the catalog in the module
    #: docstring.
    default_code: str = "DOCX_ERROR"

    def __init__(
        self,
        message: str,
        *,
        code: Optional[str] = None,
        suggestion: Optional[str] = None,
        location: Optional[str] = None,
        operation: Optional[str] = None,
    ) -> None:
        super().__init__(message)
        self.message: str = message
        self.code: str = code if code is not None else self.default_code
        self.suggestion: Optional[str] = suggestion
        self.location: Optional[str] = location
        self.operation: Optional[str] = operation

    def __str__(self) -> str:  # pragma: no cover - cosmetic
        parts = [f"[{self.code}] {self.message}"]
        if self.location:
            parts.append(f"location={self.location}")
        if self.operation:
            parts.append(f"operation={self.operation}")
        if self.suggestion:
            parts.append(f"suggestion={self.suggestion}")
        return " | ".join(parts)

    def to_dict(self) -> dict:
        """Return a plain ``dict`` view of the structured fields.

        Useful for JSON-serialising the error for an LLM-driven repair
        loop or a cross-process error log.
        """
        return {
            "code": self.code,
            "message": self.message,
            "suggestion": self.suggestion,
            "location": self.location,
            "operation": self.operation,
        }


# -- backwards-compatible aliases for the legacy exceptions -----------------


class InvalidSpanError(PythonDocxError):
    """Raised when an invalid merge region is specified in a request to merge table
    cells."""


class InvalidXmlError(PythonDocxError):
    """Raised when invalid XML is encountered, such as on attempt to access a missing
    required child element."""


class EncryptedDocumentError(PythonDocxError):
    """Raised when attempting to open a password-encrypted .docx file.

    Word stores encrypted documents as OLE compound files (CFBF) containing the
    encrypted package, which cannot be opened by the standard zipfile reader.
    Detection is performed by checking the file's magic bytes against the OLE
    compound file signature ``D0 CF 11 E0 A1 B1 1A E1``.

    Also raised when the optional ``python-ooxml-crypto`` dependency is required
    to decrypt or encrypt a package but is not installed, when the supplied
    password does not match the one used to encrypt the package, or when the
    underlying encryption container is malformed.
    """


class NestedSectionError(PythonDocxError):
    """Raised when entering a section context inside another active one.

    The OOXML model encodes sections by attaching a ``w:sectPr`` to the
    last paragraph of a region. Sections cannot nest — every paragraph
    belongs to exactly one section. :meth:`docx.Document.section`
    surfaces this constraint at the API layer.

    .. versionadded:: 2026.05.13
    """


class RmsProtectedDocumentError(EncryptedDocumentError):
    """Raised when opening a .docx wrapped in Azure RMS / AIP / IRM protection.

    "Rights Management Services" (also marketed as Azure Information Protection /
    Microsoft Purview Information Protection / "Information Rights Management")
    wraps the regular OOXML zip inside a CFBF (OLE2 compound file) container
    that stores the encrypted payload under a ``DRMContent`` stream and a
    ``DRMEncryptedTransform`` descriptor. Unlike an ECMA-376 Agile-Encryption
    package, an RMS package cannot be decrypted with a password alone — the
    user's Azure AD / Microsoft 365 identity must be presented to the RMS
    service to retrieve the content key.

    python-docx does not bundle an RMS client (the Microsoft Information
    Protection SDK is C#/.NET-only and requires an interactive Azure AD login
    flow). Callers that need RMS decryption should delegate to Microsoft Office
    automation, the MIP SDK, or a pre-processing step before opening the file
    with python-docx.

    .. versionadded:: 2026.05.10
    """


# -- structured DocxError subclasses ----------------------------------------


class StyleNotFoundError(DocxError, KeyError):
    """A style lookup by name (or ``WD_BUILTIN_STYLE``) failed.

    Multi-inherits from :class:`KeyError` so legacy
    ``except KeyError:`` callers continue to catch it.
    """

    default_code = "STYLE_NOT_FOUND"

    def __str__(self) -> str:  # pragma: no cover - cosmetic
        return DocxError.__str__(self)


class StyleDuplicateError(DocxError, ValueError):
    """Attempted to add a style whose name is already present in the document."""

    default_code = "STYLE_DUPLICATE"


class StyleTypeMismatchError(DocxError, ValueError):
    """A style is the wrong ``WD_STYLE_TYPE`` for the requested operation."""

    default_code = "STYLE_TYPE_MISMATCH"


class LatentStyleNotFoundError(DocxError, KeyError):
    """A latent-style lookup by name failed."""

    default_code = "LATENT_STYLE_NOT_FOUND"

    def __str__(self) -> str:  # pragma: no cover - cosmetic
        return DocxError.__str__(self)


class BuiltinStyleNotFoundError(DocxError, KeyError):
    """``Styles.import_builtin`` did not find the named built-in style."""

    default_code = "BUILTIN_STYLE_NOT_FOUND"

    def __str__(self) -> str:  # pragma: no cover - cosmetic
        return DocxError.__str__(self)


class BookmarkNotFoundError(DocxError, KeyError):
    """A bookmark lookup or removal by name failed."""

    default_code = "BOOKMARK_NOT_FOUND"

    def __str__(self) -> str:  # pragma: no cover - cosmetic
        return DocxError.__str__(self)


class FontNotFoundError(DocxError, KeyError):
    """``FontTable[name]`` lookup failed."""

    default_code = "FONT_NOT_FOUND"

    def __str__(self) -> str:  # pragma: no cover - cosmetic
        return DocxError.__str__(self)


class FontFamilyInvalidError(DocxError, ValueError):
    """``FontTable.add_embedded_font`` was given an unknown font-family token."""

    default_code = "FONT_FAMILY_INVALID"


class FontEmbedEmptyError(DocxError, ValueError):
    """``FontTable.embed_font`` was called with no variant bytes."""

    default_code = "FONT_EMBED_EMPTY"


class ThemeTokenInvalidError(DocxError, KeyError):
    """``ThemeColors[token]`` was given a token outside the twelve scheme slots."""

    default_code = "THEME_TOKEN_INVALID"

    def __str__(self) -> str:  # pragma: no cover - cosmetic
        return DocxError.__str__(self)


class InvalidColorError(DocxError, ValueError):
    """An RGB component or hex string did not satisfy the color contract."""

    default_code = "INVALID_COLOR"


class InvalidBrightnessError(DocxError, ValueError):
    """Theme-color brightness was out of range or no theme color was set."""

    default_code = "INVALID_BRIGHTNESS"


class OutOfRangeError(DocxError, IndexError):
    """A numeric index was out of range for the targeted collection.

    Multi-inherits from :class:`IndexError` so existing
    ``except IndexError:`` callers continue to catch it.
    """

    default_code = "OUT_OF_RANGE"

    def __str__(self) -> str:  # pragma: no cover - cosmetic
        return DocxError.__str__(self)


class ValueOutOfRangeError(DocxError, ValueError):
    """A scalar value was outside the documented permitted range.

    Distinct from :class:`OutOfRangeError` (which extends ``IndexError``)
    because the value here is a *parameter*, not a sequence index — the
    operation should fail with the legacy ``ValueError`` semantics.
    """

    default_code = "VALUE_OUT_OF_RANGE"


class NotAWordFileError(DocxError, ValueError):
    """The file's content type is not a recognised WordprocessingML document."""

    default_code = "NOT_A_WORD_FILE"


class NotAWordTemplateError(DocxError, ValueError):
    """The package is not a ``.dotx`` / ``.dotm`` Word template."""

    default_code = "NOT_A_WORD_TEMPLATE"


# -- helper utilities -------------------------------------------------------


_T = TypeVar("_T")


def closest_names(
    name: str,
    candidates: Iterable[str],
    *,
    n: int = 3,
    cutoff: float = 0.6,
) -> list[str]:
    """Return up to `n` strings from `candidates` closest to `name`.

    Thin wrapper over :func:`difflib.get_close_matches` configured with
    a slightly forgiving cutoff so single-character typos surface
    matches. The returned list is ordered best-first; an empty list
    means no close match was found.
    """
    return difflib.get_close_matches(name, list(candidates), n=n, cutoff=cutoff)


def _did_you_mean(
    name: str,
    candidates: Iterable[str],
    *,
    sample_limit: int = 5,
) -> Optional[str]:
    """Return a "did you mean…" suggestion string, or |None|.

    Builds a stable, LLM-friendly string of the form
    ``"Did you mean 'X' or 'Y'? Available: A, B, C, …"`` from a fuzzy
    match plus a sample of all available names (so the LLM has both a
    targeted hint and a way to enumerate alternatives without an
    extra introspection round-trip).
    """
    available = list(candidates)
    matches = closest_names(name, available)
    sample = ", ".join(available[:sample_limit])
    if not available:
        return None
    if matches:
        suggestion = "Did you mean " + " or ".join(repr(m) for m in matches) + "?"
        if sample:
            suggestion += f" Available: {sample}"
            if len(available) > sample_limit:
                suggestion += f", … ({len(available)} total)"
        return suggestion
    if sample:
        more = (
            f", … ({len(available)} total)" if len(available) > sample_limit else ""
        )
        return f"Available: {sample}{more}"
    return None


def _err_class_for(
    builtin: Type[BaseException],
    structured_class: Type[DocxError],
) -> Type[DocxError]:
    """Internal helper: identity passthrough used by tests and tooling.

    Returned class is `structured_class`. The signature mirrors what a
    future "auto-pick a base from `builtin`" helper might look like;
    today this is simply a typed lookup.
    """
    del builtin
    return structured_class


# -- public re-exports for convenient catch tuples -------------------------

#: Catch-all tuple every "name not found" scenario satisfies. Useful for
#: callers writing ``except NameNotFoundError`` style guards.
NameNotFoundError: Tuple[Type[DocxError], ...] = (
    StyleNotFoundError,
    LatentStyleNotFoundError,
    BuiltinStyleNotFoundError,
    BookmarkNotFoundError,
    FontNotFoundError,
    ThemeTokenInvalidError,
)


def is_docx_error(exc: BaseException) -> bool:
    """Return True if `exc` is a structured :class:`DocxError`.

    Convenience for log filters that want to special-case structured
    errors without importing every subclass.
    """
    return isinstance(exc, DocxError)


# -- ensure ``_T`` and ``_err_class_for`` are referenced so type-checkers --
# -- don't complain about unused symbols in the wheel. --
_ = (_T, _err_class_for, Union)
