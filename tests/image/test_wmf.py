"""Unit test suite for docx.image.wmf module."""

from __future__ import annotations

import io
import struct

import pytest

from docx.image.constants import MIME_TYPE
from docx.image.exceptions import InvalidImageStreamError
from docx.image.image import _ImageHeaderFactory
from docx.image.wmf import Wmf


def _placeable_header(
    bounds=(0, 0, 1440, 720),
    inch=1440,
):
    """Build a minimal Aldus placeable WMF header (22 bytes)."""
    header = bytearray(22)
    struct.pack_into("<I", header, 0x00, 0x9AC6CDD7)   # Key
    struct.pack_into("<H", header, 0x04, 0)            # HWmf
    struct.pack_into("<hhhh", header, 0x06, *bounds)   # BoundingBox
    struct.pack_into("<H", header, 0x0E, inch)         # Inch
    struct.pack_into("<I", header, 0x10, 0)            # Reserved
    struct.pack_into("<H", header, 0x14, 0)            # Checksum
    return bytes(header)


class DescribeWmf:
    def it_parses_a_placeable_wmf_header(self):
        blob = _placeable_header(bounds=(0, 0, 1440, 720), inch=1440)
        stream = io.BytesIO(blob)

        wmf = Wmf.from_stream(stream)

        # 1440 units at 1440 units/inch = 1 inch = 96 px at 96 dpi
        assert wmf.px_width == 96
        assert wmf.px_height == 48
        assert wmf.horz_dpi == 1440
        assert wmf.vert_dpi == 1440

    def it_knows_its_content_type(self):
        assert Wmf(0, 0, 0, 0).content_type == MIME_TYPE.WMF

    def it_knows_its_default_ext(self):
        assert Wmf(0, 0, 0, 0).default_ext == "wmf"

    def it_rejects_a_non_placeable_stream(self):
        blob = b"\x01\x00\x09\x00" + b"\x00" * 40
        with pytest.raises(InvalidImageStreamError):
            Wmf.from_stream(io.BytesIO(blob))

    def it_rejects_a_placeable_header_with_zero_inch(self):
        blob = _placeable_header(inch=0)
        with pytest.raises(InvalidImageStreamError):
            Wmf.from_stream(io.BytesIO(blob))

    def it_handles_negative_bounds(self):
        blob = _placeable_header(bounds=(-100, -50, 1340, 670), inch=1440)
        wmf = Wmf.from_stream(io.BytesIO(blob))
        # 1440 units wide, 720 tall → 96 x 48 at 96 dpi
        assert wmf.px_width == 96
        assert wmf.px_height == 48

    def it_is_recognized_by_the_image_factory(self):
        blob = _placeable_header()
        stream = io.BytesIO(blob)

        header = _ImageHeaderFactory(stream)

        assert isinstance(header, Wmf)
