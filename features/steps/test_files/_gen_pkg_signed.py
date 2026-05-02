"""Generate ``pkg-signed.docx`` fixture for digital-signature detection scenarios.

python-docx cannot create, verify, or counter-sign digital signatures, but it
*does* surface them: :attr:`docx.document.Document.is_signed` reports presence
and :attr:`docx.document.Document.signatures` exposes one
:class:`.SignatureInfo` per signature part, with ``signer`` and ``signed_at``
parsed from the embedded XML-DSig / XAdES blobs.

This generator builds a minimal signed package by hand from the default
python-docx template:

1. Adds a package-level ``digital-signature/origin`` relationship targeting
   ``/_xmlsignatures/origin.sigs``.
2. Writes a placeholder binary origin part (python-docx does not parse it).
3. Writes an ``origin.sigs.rels`` relationship file referencing
   ``sig1.xml`` with reltype ``digital-signature/signature``.
4. Writes a minimal XML-DSig ``sig1.xml`` carrying an ``X509SubjectName`` of
   ``CN=Alice Example`` and a XAdES ``SigningTime`` of ``2024-04-01T12:34:56Z``.

The signature blob is *not* cryptographically valid — Word would reject it —
but python-docx's detection and metadata-extraction logic treats it as a
signed package and so is usable for behave coverage.

Run ``python features/steps/test_files/_gen_pkg_signed.py`` to regenerate.
"""

from __future__ import annotations

import datetime as dt
import os
import zipfile

from docx import Document

HERE = os.path.abspath(os.path.dirname(__file__))
SOURCE = os.path.normpath(
    os.path.join(HERE, "..", "..", "..", "src", "docx", "templates", "default.docx")
)
OUT_PATH = os.path.join(HERE, "pkg-signed.docx")

_SIG_ORIGIN_CT = "application/vnd.openxmlformats-package.digital-signature-origin"
_SIG_XML_CT = "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"
_SIG_ORIGIN_RT = (
    "http://schemas.openxmlformats.org/package/2006/relationships/"
    "digital-signature/origin"
)
_SIG_RT = (
    "http://schemas.openxmlformats.org/package/2006/relationships/"
    "digital-signature/signature"
)

_SIG_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Signature xmlns="http://www.w3.org/2000/09/xmldsig#">
  <SignedInfo/>
  <SignatureValue>ignored</SignatureValue>
  <KeyInfo>
    <X509Data>
      <X509SubjectName>CN=Alice Example</X509SubjectName>
    </X509Data>
  </KeyInfo>
  <Object>
    <SignatureProperties xmlns:xades="http://uri.etsi.org/01903/v1.3.2#">
      <SignatureProperty>
        <xades:SigningTime>2024-04-01T12:34:56Z</xades:SigningTime>
      </SignatureProperty>
    </SignatureProperties>
  </Object>
</Signature>
"""

_ORIGIN_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rIdSig1" Type="{_SIG_RT}" Target="sig1.xml"/>'
    "</Relationships>"
).encode("utf-8")


def _patch_content_types(xml: bytes) -> bytes:
    overrides = (
        f'<Override PartName="/_xmlsignatures/origin.sigs"'
        f' ContentType="{_SIG_ORIGIN_CT}"/>'
        f'<Override PartName="/_xmlsignatures/sig1.xml"'
        f' ContentType="{_SIG_XML_CT}"/>'
    ).encode("utf-8")
    end = xml.rfind(b"</Types>")
    if end == -1:
        raise ValueError("[Content_Types].xml missing </Types>")
    return xml[:end] + overrides + xml[end:]


def _patch_package_rels(xml: bytes) -> bytes:
    rel = (
        f'<Relationship Id="rIdSigOrigin" Type="{_SIG_ORIGIN_RT}"'
        f' Target="/_xmlsignatures/origin.sigs"/>'
    ).encode("utf-8")
    end = xml.rfind(b"</Relationships>")
    if end == -1:
        raise ValueError("_rels/.rels missing </Relationships>")
    return xml[:end] + rel + xml[end:]


def build() -> str:
    if not os.path.isfile(SOURCE):
        raise FileNotFoundError(SOURCE)

    with zipfile.ZipFile(SOURCE, "r") as zi, zipfile.ZipFile(
        OUT_PATH, "w", zipfile.ZIP_DEFLATED
    ) as zo:
        for info in zi.infolist():
            data = zi.read(info.filename)
            if info.filename == "[Content_Types].xml":
                data = _patch_content_types(data)
            elif info.filename == "_rels/.rels":
                data = _patch_package_rels(data)
            zo.writestr(info, data)
        zo.writestr("_xmlsignatures/origin.sigs", b"")
        zo.writestr("_xmlsignatures/_rels/origin.sigs.rels", _ORIGIN_RELS)
        zo.writestr("_xmlsignatures/sig1.xml", _SIG_XML)

    return OUT_PATH


def validate(path: str) -> None:
    document = Document(path)
    assert document.is_signed is True, "is_signed was False for signed package"
    sigs = document.signatures
    assert len(sigs) == 1, f"expected 1 signature, got {len(sigs)}"
    sig = sigs[0]
    assert sig.signer == "CN=Alice Example", sig.signer
    assert sig.signed_at == dt.datetime(
        2024, 4, 1, 12, 34, 56, tzinfo=dt.timezone.utc
    ), sig.signed_at
    assert str(sig.partname) == "/_xmlsignatures/sig1.xml"


if __name__ == "__main__":
    out = build()
    validate(out)
    print(f"wrote {out}")
