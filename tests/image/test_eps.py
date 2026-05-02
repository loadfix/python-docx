"""Unit test suite for docx.image.eps module."""

from __future__ import annotations

import io
import struct

import pytest

from docx.image.constants import MIME_TYPE
from docx.image.eps import Eps, is_eps_stream
from docx.image.exceptions import InvalidImageStreamError
from docx.image.image import _ImageHeaderFactory


_PLAIN_EPS = (
    b"%!PS-Adobe-3.0 EPSF-3.0\n"
    b"%%BoundingBox: 0 0 612 792\n"
    b"%%Creator: test\n"
    b"%%EndComments\n"
    b"showpage\n"
)


class DescribeEps:
    def it_parses_a_plain_eps_stream(self):
        stream = io.BytesIO(_PLAIN_EPS)

        eps = Eps.from_stream(stream)

        assert eps.px_width == 612
        assert eps.px_height == 792
        assert eps.horz_dpi == 72
        assert eps.vert_dpi == 72

    def it_parses_a_dos_eps_stream(self):
        ps = _PLAIN_EPS
        ps_offset = 30
        dos_header = (
            b"\xc5\xd0\xd3\xc6"
            + struct.pack("<II", ps_offset, len(ps))
            + b"\x00" * (ps_offset - 12)
        )
        blob = dos_header + ps

        eps = Eps.from_stream(io.BytesIO(blob))

        assert eps.px_width == 612
        assert eps.px_height == 792

    def it_knows_its_content_type(self):
        assert Eps(0, 0, 0, 0).content_type == MIME_TYPE.EPS

    def it_knows_its_default_ext(self):
        assert Eps(0, 0, 0, 0).default_ext == "eps"

    def it_rejects_non_eps_text(self):
        with pytest.raises(InvalidImageStreamError):
            Eps.from_stream(io.BytesIO(b"not postscript"))

    def it_rejects_eps_without_bounding_box(self):
        blob = b"%!PS-Adobe-3.0 EPSF-3.0\nshowpage\n"
        with pytest.raises(InvalidImageStreamError):
            Eps.from_stream(io.BytesIO(blob))

    def it_parses_negative_bounding_box(self):
        blob = (
            b"%!PS-Adobe-3.0 EPSF-3.0\n"
            b"%%BoundingBox: -10 -20 100 80\n"
            b"showpage\n"
        )
        eps = Eps.from_stream(io.BytesIO(blob))
        assert eps.px_width == 110
        assert eps.px_height == 100

    def it_is_recognized_by_the_image_factory(self):
        stream = io.BytesIO(_PLAIN_EPS)

        header = _ImageHeaderFactory(stream)

        assert isinstance(header, Eps)


class DescribeIsEpsStream:
    def it_detects_plain_eps(self):
        assert is_eps_stream(io.BytesIO(_PLAIN_EPS)) is True

    def it_detects_dos_eps(self):
        assert is_eps_stream(io.BytesIO(b"\xc5\xd0\xd3\xc6" + b"\x00" * 60)) is True

    def it_rejects_non_eps(self):
        assert is_eps_stream(io.BytesIO(b"random bytes")) is False
