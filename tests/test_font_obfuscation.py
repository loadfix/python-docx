"""Unit-test suite for `docx.font_obfuscation`.

Covers the ECMA-376 Part 1 §17.8 XOR obfuscation used by Word to store
embedded TrueType bytes under the
``application/vnd.openxmlformats-officedocument.obfuscatedFont``
content-type.
"""

from __future__ import annotations

import os

import pytest

from docx.font_obfuscation import (
    OBFUSCATED_FONT_CONTENT_TYPE,
    deobfuscate_font_bytes,
    derive_obfuscation_key,
    generate_font_key,
    obfuscate_font_bytes,
)


class DescribeFontObfuscation:
    """End-to-end checks for the obfuscation/deobfuscation helpers."""

    def it_declares_the_obfuscated_font_content_type(self):
        assert OBFUSCATED_FONT_CONTENT_TYPE == (
            "application/vnd.openxmlformats-officedocument.obfuscatedFont"
        )

    def it_generates_a_fresh_GUID_each_call(self):
        k1 = generate_font_key()
        k2 = generate_font_key()
        assert k1 != k2
        # -- canonical braced uppercase form, 38 chars incl. braces --
        assert k1.startswith("{") and k1.endswith("}")
        assert len(k1) == 38

    def it_derives_a_16_byte_key_from_the_fontKey_GUID(self):
        # -- a contrived "00112233-4455-6677-8899-AABBCCDDEEFF" GUID has the
        # -- raw hex 001122334455667788 99AABBCCDDEEFF; reversing gives the
        # -- expected obfuscation key. --
        guid = "{00112233-4455-6677-8899-AABBCCDDEEFF}"
        key = derive_obfuscation_key(guid)
        assert key == bytes.fromhex("FFEEDDCCBBAA99887766554433221100")

    def it_round_trips_a_typical_font_blob(self):
        # -- give the fake font bytes a recognisable pattern so the test
        # -- can distinguish "XOR applied" from "XOR not applied". --
        font_bytes = bytes(range(200))
        guid = generate_font_key()

        obfuscated = obfuscate_font_bytes(font_bytes, guid)

        # -- bytes beyond offset 32 are unchanged --
        assert obfuscated[32:] == font_bytes[32:]
        # -- at least one of the first 32 bytes must have changed --
        assert obfuscated[:32] != font_bytes[:32]

        restored = deobfuscate_font_bytes(obfuscated, guid)
        assert restored == font_bytes

    def it_handles_blobs_shorter_than_32_bytes(self):
        font_bytes = b"\x01\x02\x03\x04\x05"
        guid = generate_font_key()

        obfuscated = obfuscate_font_bytes(font_bytes, guid)
        restored = deobfuscate_font_bytes(obfuscated, guid)

        assert len(obfuscated) == len(font_bytes)
        assert restored == font_bytes

    def it_is_symmetric_obfuscating_twice_restores_the_bytes(self):
        font_bytes = os.urandom(256)
        guid = generate_font_key()

        once = obfuscate_font_bytes(font_bytes, guid)
        twice = obfuscate_font_bytes(once, guid)

        assert twice == font_bytes

    def it_accepts_both_braced_and_unbraced_GUIDs(self):
        font_bytes = b"A" * 64
        braced = "{AABBCCDD-EEFF-0011-2233-445566778899}"
        unbraced = "AABBCCDD-EEFF-0011-2233-445566778899"

        assert obfuscate_font_bytes(font_bytes, braced) == obfuscate_font_bytes(
            font_bytes, unbraced
        )

    @pytest.mark.parametrize(
        "bad_guid",
        [
            "",
            "not-a-guid",
            "{0011}",
            "{ZZZZZZZZ-4455-6677-8899-AABBCCDDEEFF}",
        ],
    )
    def it_rejects_malformed_GUIDs(self, bad_guid: str):
        with pytest.raises(ValueError):
            derive_obfuscation_key(bad_guid)
