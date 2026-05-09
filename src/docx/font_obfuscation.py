"""Pure-Python ECMA-376 Part 1 §17.8 obfuscated-font round-trip.

The .docx "embedded font" mechanism stores raw TrueType bytes as a
package part with the content-type
``application/vnd.openxmlformats-officedocument.obfuscatedFont``. Before
the bytes are written the first 32 bytes are XOR-obfuscated using a key
derived from the ``w:fontKey`` GUID attached to the
``<w:embedRegular>``/``<w:embedBold>``/``<w:embedItalic>``/
``<w:embedBoldItalic>`` child. The operation is symmetric (XOR) so the
same routine handles deobfuscation.

The key derivation follows §17.8:

1. Strip braces and hyphens from the GUID string and parse the remaining
   32 hex characters as 16 raw bytes.
2. Reverse the byte order. The resulting 16-byte sequence is the
   *obfuscation key*.
3. XOR the first 32 bytes of the font data with the key: bytes 0..15 are
   XORed with the key, and bytes 16..31 are XORed with the key again.
   Remaining bytes pass through unchanged.

This module is intentionally dependency-free — no cryptography package
is required — because the obfuscation is a simple per-byte XOR.

.. versionadded:: 2026.05.10
"""

from __future__ import annotations

import uuid
from typing import Union

__all__ = (
    "OBFUSCATED_FONT_CONTENT_TYPE",
    "derive_obfuscation_key",
    "generate_font_key",
    "obfuscate_font_bytes",
    "deobfuscate_font_bytes",
)


#: OPC content-type used for embedded obfuscated-font parts.
OBFUSCATED_FONT_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.obfuscatedFont"
)


def generate_font_key() -> str:
    """Return a fresh fontKey GUID formatted as ``{XXXXXXXX-XXXX-...}``.

    The casing and brace-wrapping match what Word emits, which in turn
    matches ST_Guid from the shared schema. Each call returns a
    version-4 random UUID.
    """
    return "{" + str(uuid.uuid4()).upper() + "}"


def _normalise_guid(guid: str) -> bytes:
    """Return the 16 raw bytes of `guid` (braces and hyphens optional).

    Accepts the casing and punctuation that Word writes for ``w:fontKey``
    (``{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}``) and the unbraced form
    returned by :class:`uuid.UUID`.
    """
    stripped = guid.strip().lstrip("{").rstrip("}").replace("-", "")
    if len(stripped) != 32:
        raise ValueError(
            f"fontKey must contain 32 hex characters (excluding braces "
            f"and hyphens), got {len(stripped)}: {guid!r}"
        )
    try:
        return bytes.fromhex(stripped)
    except ValueError as exc:
        raise ValueError(f"fontKey is not a valid hex GUID: {guid!r}") from exc


def derive_obfuscation_key(font_key: str) -> bytes:
    """Return the 16-byte obfuscation key derived from `font_key`.

    Implements the reversal rule of ECMA-376 Part 1 §17.8: the raw
    GUID bytes are taken in reverse order to form the XOR key.
    """
    return _normalise_guid(font_key)[::-1]


def _xor_first_32_bytes(blob: bytes, font_key: str) -> bytes:
    """Return `blob` with its first 32 bytes XORed against the derived key.

    XOR is an involution, so a second pass over an obfuscated blob
    restores the original bytes. The function tolerates blobs shorter
    than 32 bytes (it XORs whatever is available and passes the rest
    through untouched).
    """
    key = derive_obfuscation_key(font_key)
    # -- copy so we can mutate --
    data = bytearray(blob)
    prefix_len = min(32, len(data))
    for i in range(prefix_len):
        data[i] ^= key[i % 16]
    return bytes(data)


def obfuscate_font_bytes(font_bytes: Union[bytes, bytearray], font_key: str) -> bytes:
    """Return `font_bytes` with the ECMA-376 §17.8 XOR obfuscation applied.

    `font_bytes` is the raw (un-obfuscated) font file as loaded from disk
    or memory, and `font_key` is the GUID value that will be written to
    the ``w:fontKey`` attribute pointing at this part. The result is the
    byte sequence that must be stored in the package part.
    """
    return _xor_first_32_bytes(bytes(font_bytes), font_key)


def deobfuscate_font_bytes(
    obfuscated: Union[bytes, bytearray], font_key: str
) -> bytes:
    """Return the original font bytes given the `obfuscated` package bytes.

    XOR is symmetric — this is identical to
    :func:`obfuscate_font_bytes`, but exposed under its semantic name so
    callers don't have to read the algorithm comments to understand which
    direction a given call is going in.
    """
    return _xor_first_32_bytes(bytes(obfuscated), font_key)
