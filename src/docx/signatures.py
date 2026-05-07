"""Digital-signature detection helpers.

python-docx does not create or verify XML-DSig / XAdES signatures, but it does surface
whether a ``.docx`` package contains digital-signature parts and exposes minimal metadata
parsed from those parts so callers can decide what to do.

A signed OOXML package contains:

* a package-level relationship of type
  ``.../digital-signature/origin`` that targets an ``_xmlsignatures/origin.sigs`` part;
* one or more relationships of type ``.../digital-signature/signature`` from the origin
  part, each targeting a ``/_xmlsignatures/sigN.xml`` part holding an XML-DSig document,
  optionally with XAdES extensions carrying the signing time and signer identity.

The shared-package integration (``python-ooxml-signatures`` / ``ooxml_signatures``)
is picked up opportunistically: when that package is importable,
:attr:`SignatureInfo.shared_signature` returns the corresponding
``ooxml_signatures.Signature`` instance and the metadata accessors
(``signer``, ``signed_at``) delegate to the richer shared-package parser,
which supports Microsoft's ``mdssi:SignatureTime`` + ``mdssi:SignatureComments``
extensions alongside XAdES. When the shared package is not installed, the
legacy inline parser continues to handle the common happy path unchanged.
"""

from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING, Any

from docx.oxml.parser import parse_xml

if TYPE_CHECKING:
    from docx.opc.packuri import PackURI
    from docx.opc.part import Part


_XMLDSIG_NS = "http://www.w3.org/2000/09/xmldsig#"
_XADES_NS = "http://uri.etsi.org/01903/v1.3.2#"


def _import_ooxml_signatures() -> Any:
    """Return the ``ooxml_signatures`` module, or |None| if not installed.

    Kept as a function (rather than a module-level import) so that the
    import is attempted lazily on first access. Tests can monkey-patch
    this symbol to force the fallback path regardless of the real
    environment.
    """
    try:
        import ooxml_signatures  # type: ignore[import-not-found]

        return ooxml_signatures
    except ImportError:
        return None


class SignatureInfo:
    """Read-only metadata for a single digital signature in a package.

    Instances are produced by :attr:`docx.document.Document.signatures`; they are not
    intended to be constructed directly by library users.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, part: Part):
        self._part = part
        self._parsed: tuple[str | None, datetime | None] | None = None
        self._shared: Any = None
        self._shared_resolved = False

    @property
    def partname(self) -> PackURI:
        """The OPC part name, e.g. ``/_xmlsignatures/sig1.xml``.

        .. versionadded:: 2026.05.0
        """
        return self._part.partname

    @property
    def blob(self) -> bytes:
        """Raw XML bytes of the signature part.

        .. versionadded:: 2026.05.0
        """
        return self._part.blob

    @property
    def shared_signature(self) -> Any:
        """Return the ``ooxml_signatures.Signature`` for this part, or |None|.

        Present when ``python-ooxml-signatures`` is installed; |None| otherwise.
        The richer shared-package parser supports Microsoft's
        ``mdssi:SignatureTime`` and ``mdssi:SignatureComments`` extensions
        alongside XAdES ``SigningTime``, and exposes ``references`` /
        ``comments`` attributes that python-docx's inline parser doesn't.

        .. versionadded:: 2026.05.0
        """
        if not self._shared_resolved:
            mod = _import_ooxml_signatures()
            if mod is not None:
                try:
                    self._shared = mod.Signature.from_bytes(
                        self.blob, partname=str(self.partname)
                    )
                except Exception:  # noqa: BLE001 — malformed → no shared proxy
                    self._shared = None
            self._shared_resolved = True
        return self._shared

    @property
    def signer(self) -> str | None:
        """Subject name of the signing certificate, or |None| if not present.

        Extracted from ``<X509SubjectName>`` under the XML-DSig ``KeyInfo`` or from the
        XAdES ``SigningCertificate`` block. Returns |None| when the signature XML is
        malformed or does not expose this information. When
        ``python-ooxml-signatures`` is installed, this delegates to
        :attr:`ooxml_signatures.Signature.signer` for richer handling
        (picks up non-default ``X509SubjectName`` locations that Office
        emits for XAdES signatures).

        .. versionadded:: 2026.05.0
        """
        shared = self.shared_signature
        if shared is not None:
            return shared.signer
        return self._parse()[0]

    @property
    def signed_at(self) -> datetime | None:
        """Time the signature was created, or |None| if not declared.

        Parsed from the XAdES ``<SigningTime>`` element when present. Returns |None|
        when the signature XML is malformed or does not declare a signing time. When
        ``python-ooxml-signatures`` is installed, this delegates to
        :attr:`ooxml_signatures.Signature.signed_at` which prefers
        Microsoft's ``mdssi:SignatureTime`` (the shape Office writes by
        default) and falls back to XAdES — the inline parser only sees
        the XAdES case.

        .. versionadded:: 2026.05.0
        """
        shared = self.shared_signature
        if shared is not None:
            return shared.signed_at
        return self._parse()[1]

    def _parse(self) -> tuple[str | None, datetime | None]:
        if self._parsed is None:
            self._parsed = _parse_signature_xml(self.blob)
        return self._parsed


def _parse_signature_xml(blob: bytes) -> tuple[str | None, datetime | None]:
    """Return ``(signer, signed_at)`` parsed from signature-part XML `blob`.

    Returns ``(None, None)`` when `blob` cannot be parsed or the expected elements are
    absent. Callers therefore do not need to handle exceptions from malformed input.
    """
    if not blob:
        return (None, None)
    try:
        root = parse_xml(blob)
    except Exception:
        return (None, None)

    signer = _extract_signer(root)
    signed_at = _extract_signed_at(root)
    return (signer, signed_at)


def _extract_signer(root: Any) -> str | None:
    """Return the X509 subject name from `root`, or |None| if not found."""
    # -- XML-DSig: <Signature>/<KeyInfo>/<X509Data>/<X509SubjectName>; accept a bare
    # -- <X509SubjectName> anywhere under root for robustness. --
    tag = f"{{{_XMLDSIG_NS}}}X509SubjectName"
    try:
        # lxml Element.iter() accepts Clark-notation tag names
        for elem in root.iter(tag):
            text = elem.text
            if text:
                return str(text).strip()
    except Exception:
        return None
    return None


def _extract_signed_at(root: Any) -> datetime | None:
    """Return XAdES ``SigningTime`` parsed as a |datetime|, or |None| if absent."""
    tag = f"{{{_XADES_NS}}}SigningTime"
    try:
        for elem in root.iter(tag):
            text = elem.text
            if not text:
                continue
            parsed = _parse_iso_datetime(str(text).strip())
            if parsed is not None:
                return parsed
    except Exception:
        return None
    return None


def _parse_iso_datetime(text: str) -> datetime | None:
    """Best-effort ISO-8601 parser for XAdES ``SigningTime`` values.

    XAdES signing times are ``xsd:dateTime`` which may carry a ``Z`` suffix. Python's
    :func:`datetime.fromisoformat` handles the ``Z`` suffix from 3.11+; for older
    interpreters we fall back to replacing the ``Z`` with a ``+00:00`` offset.
    """
    try:
        return datetime.fromisoformat(text)
    except ValueError:
        pass
    if text.endswith("Z"):
        try:
            return datetime.fromisoformat(text[:-1] + "+00:00")
        except ValueError:
            return None
    return None
