"""Unit test suite for docx.image.webp module."""

from __future__ import annotations

import io
import struct

import pytest

from docx.image.constants import MIME_TYPE
from docx.image.exceptions import InvalidImageStreamError
from docx.image.image import _ImageHeaderFactory
from docx.image.webp import WebP


def _riff_container(chunk_fourcc: bytes, chunk_payload: bytes) -> bytes:
    """Build a minimal RIFF/WEBP container wrapping `chunk_fourcc`+payload."""
    chunk = chunk_fourcc + struct.pack("<I", len(chunk_payload)) + chunk_payload
    inner = b"WEBP" + chunk
    return b"RIFF" + struct.pack("<I", len(inner)) + inner


def _vp8_payload(width: int, height: int) -> bytes:
    """Build a minimal VP8 lossy payload with the given 14-bit dims."""
    # 3-byte frame tag (keyframe, bit0=0) + 3-byte start code + 2-byte width
    # with 2-bit scale + 2-byte height with 2-bit scale. Remaining bytes are
    # bitstream which we never parse.
    frame_tag = b"\x00\x00\x00"
    start_code = b"\x9d\x01\x2a"
    w = struct.pack("<H", width & 0x3FFF)
    h = struct.pack("<H", height & 0x3FFF)
    return frame_tag + start_code + w + h + b"\x00" * 8


def _vp8l_payload(width: int, height: int) -> bytes:
    """Build a minimal VP8L lossless payload encoding (w-1, h-1)."""
    sig = b"\x2f"
    bits = ((width - 1) & 0x3FFF) | (((height - 1) & 0x3FFF) << 14)
    return sig + struct.pack("<I", bits)


def _vp8x_payload(width: int, height: int) -> bytes:
    """Build a minimal VP8X extended payload encoding (w-1, h-1) as 24-bit."""
    flags = b"\x00" * 4  # flags + reserved
    w_minus_1 = width - 1
    h_minus_1 = height - 1
    w_bytes = bytes((w_minus_1 & 0xFF, (w_minus_1 >> 8) & 0xFF, (w_minus_1 >> 16) & 0xFF))
    h_bytes = bytes((h_minus_1 & 0xFF, (h_minus_1 >> 8) & 0xFF, (h_minus_1 >> 16) & 0xFF))
    return flags + w_bytes + h_bytes


class DescribeWebP:
    def it_parses_a_vp8_lossy_stream(self):
        blob = _riff_container(b"VP8 ", _vp8_payload(200, 100))
        stream = io.BytesIO(blob)

        webp = WebP.from_stream(stream)

        assert webp.px_width == 200
        assert webp.px_height == 100
        assert webp.horz_dpi == 72
        assert webp.vert_dpi == 72

    def it_parses_a_vp8l_lossless_stream(self):
        blob = _riff_container(b"VP8L", _vp8l_payload(640, 480))
        stream = io.BytesIO(blob)

        webp = WebP.from_stream(stream)

        assert webp.px_width == 640
        assert webp.px_height == 480

    def it_parses_a_vp8x_extended_stream(self):
        blob = _riff_container(b"VP8X", _vp8x_payload(1024, 768))
        stream = io.BytesIO(blob)

        webp = WebP.from_stream(stream)

        assert webp.px_width == 1024
        assert webp.px_height == 768

    def it_knows_its_content_type(self):
        assert WebP(0, 0, 0, 0).content_type == MIME_TYPE.WEBP

    def it_knows_its_default_ext(self):
        assert WebP(0, 0, 0, 0).default_ext == "webp"

    def it_rejects_a_non_riff_stream(self):
        with pytest.raises(InvalidImageStreamError):
            WebP.from_stream(io.BytesIO(b"NOT A WEBP" + b"\x00" * 64))

    def it_rejects_an_unknown_chunk_type(self):
        blob = b"RIFF" + struct.pack("<I", 16) + b"WEBP" + b"ZZZZ" + b"\x00" * 40
        with pytest.raises(InvalidImageStreamError):
            WebP.from_stream(io.BytesIO(blob))

    def it_is_recognized_by_the_image_factory(self):
        blob = _riff_container(b"VP8L", _vp8l_payload(32, 16))
        stream = io.BytesIO(blob)

        header = _ImageHeaderFactory(stream)

        assert isinstance(header, WebP)
        assert header.px_width == 32
        assert header.px_height == 16
