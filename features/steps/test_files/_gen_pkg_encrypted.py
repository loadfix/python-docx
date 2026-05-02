"""Generate ``pkg-encrypted.docx`` fixture for encrypted-docx detection scenarios.

python-docx cannot decrypt password-protected Word documents (Word wraps them
in an OLE compound file containing the encrypted package), but it *does*
detect the OLE compound file signature (``D0 CF 11 E0 A1 B1 1A E1``) and raise
:class:`docx.exceptions.EncryptedDocumentError` with an actionable message.

To exercise that detection path we don't need a *real* encrypted document —
only a file that begins with the OLE magic bytes. The fixture here is the
magic header followed by padding bytes, which is enough for the detector to
fire.

Run ``python features/steps/test_files/_gen_pkg_encrypted.py`` to regenerate.
"""

from __future__ import annotations

import os

from docx import Document
from docx.exceptions import EncryptedDocumentError

HERE = os.path.abspath(os.path.dirname(__file__))
OUT_PATH = os.path.join(HERE, "pkg-encrypted.docx")

_OLE_COMPOUND_FILE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def build() -> str:
    # -- Minimal OLE-compound-file stub: signature + 504 bytes of padding. We
    # -- don't emit a real CFBF structure because python-docx's detection
    # -- short-circuits at the first 8 bytes. Real files would be 4 KB+. --
    with open(OUT_PATH, "wb") as f:
        f.write(_OLE_COMPOUND_FILE_SIGNATURE)
        f.write(b"\x00" * 504)
    return OUT_PATH


def validate(path: str) -> None:
    try:
        Document(path)
    except EncryptedDocumentError:
        return
    raise AssertionError(
        "expected Document() to raise EncryptedDocumentError for OLE-stubbed file"
    )


if __name__ == "__main__":
    out = build()
    validate(out)
    print(f"wrote {out}")
