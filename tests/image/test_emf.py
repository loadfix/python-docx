"""Unit test suite for docx.image.emf module."""

from __future__ import annotations

import io
import struct

import pytest

from docx.image.constants import MIME_TYPE
from docx.image.emf import Emf
from docx.image.exceptions import InvalidImageStreamError
from docx.image.image import _ImageHeaderFactory


def _emf_header(
    bounds=(0, 0, 200, 100),
    frame=(0, 0, 5291, 2645),  # 0.01-mm units; 5291 ≈ 2 in, 2645 ≈ 1 in
    device_px=(1920, 1080),
    device_mm=(508, 285),  # 20 in x ~11.22 in → 96 dpi both axes
):
    """Build a minimal but valid EMR_HEADER record."""
    header = bytearray(88)
    # offset 0x00 RecordType = 1
    struct.pack_into("<I", header, 0x00, 1)
    # offset 0x04 RecordSize (arbitrary)
    struct.pack_into("<I", header, 0x04, 88)
    # offset 0x08 Bounds (RectL: left, top, right, bottom)
    struct.pack_into("<iiii", header, 0x08, *bounds)
    # offset 0x18 Frame (RectL)
    struct.pack_into("<iiii", header, 0x18, *frame)
    # offset 0x28 Signature = ' EMF'
    header[0x28:0x2C] = b" EMF"
    # offset 0x2C Version
    struct.pack_into("<I", header, 0x2C, 0x00010000)
    # 0x30 Size, 0x34 Handles, 0x36 Reserved, 0x38 nDescription,
    # 0x3C offDescription, 0x40 PalEntries — leave zero.
    # offset 0x44 Device (SizeL: cx, cy)
    struct.pack_into("<II", header, 0x44, *device_px)
    # offset 0x4C Millimeters (SizeL)
    struct.pack_into("<II", header, 0x4C, *device_mm)
    return bytes(header)


class DescribeEmf:
    def it_parses_a_valid_emf_header(self):
        blob = _emf_header()
        stream = io.BytesIO(blob)

        emf = Emf.from_stream(stream)

        # 5291 himetric ≈ 200 px; 2645 himetric ≈ 100 px at 96 dpi
        assert emf.px_width == 200
        assert emf.px_height == 100
        assert emf.horz_dpi == 96
        assert emf.vert_dpi == 96

    def it_knows_its_content_type(self):
        assert Emf(0, 0, 0, 0).content_type == MIME_TYPE.EMF

    def it_knows_its_default_ext(self):
        assert Emf(0, 0, 0, 0).default_ext == "emf"

    def it_rejects_wrong_record_type(self):
        blob = bytearray(_emf_header())
        struct.pack_into("<I", blob, 0x00, 99)  # not EMR_HEADER
        with pytest.raises(InvalidImageStreamError):
            Emf.from_stream(io.BytesIO(bytes(blob)))

    def it_rejects_missing_emf_signature(self):
        blob = bytearray(_emf_header())
        blob[0x28:0x2C] = b"NOPE"
        with pytest.raises(InvalidImageStreamError):
            Emf.from_stream(io.BytesIO(bytes(blob)))

    def it_falls_back_to_96_dpi_when_device_fields_are_zero(self):
        blob = _emf_header(device_px=(0, 0), device_mm=(0, 0))
        emf = Emf.from_stream(io.BytesIO(blob))
        assert emf.horz_dpi == 96
        assert emf.vert_dpi == 96

    def it_is_recognized_by_the_image_factory(self):
        blob = _emf_header()
        stream = io.BytesIO(blob)

        header = _ImageHeaderFactory(stream)

        assert isinstance(header, Emf)
