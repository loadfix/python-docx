"""Test suite for docx.image.bmp module."""

import io

import pytest

from docx.image.bmp import Bmp
from docx.image.constants import MIME_TYPE

from ..unitutil.mock import ANY, initializer_mock


class DescribeBmp:
    def it_can_construct_from_a_bmp_stream(self, Bmp__init__):
        cx, cy, horz_dpi, vert_dpi = 26, 43, 200, 96
        bytes_ = (
            b"fillerfillerfiller\x1a\x00\x00\x00\x2b\x00\x00\x00"
            b"fillerfiller\xb8\x1e\x00\x00\x00\x00\x00\x00"
        )
        stream = io.BytesIO(bytes_)

        bmp = Bmp.from_stream(stream)

        Bmp__init__.assert_called_once_with(ANY, cx, cy, horz_dpi, vert_dpi)
        assert isinstance(bmp, Bmp)

    def it_knows_its_content_type(self):
        bmp = Bmp(None, None, None, None)
        assert bmp.content_type == MIME_TYPE.BMP

    def it_knows_its_default_ext(self):
        bmp = Bmp(None, None, None, None)
        assert bmp.default_ext == "bmp"

    @pytest.mark.parametrize(
        ("px_per_meter", "expected_dpi"),
        [
            (0, 96),
            (None, 96),
            (1, 96),  # -- rounds to zero, falls back --
            (3780, 96),  # -- ~96 dpi --
            (11811, 300),
        ],
    )
    def it_falls_back_when_px_per_meter_is_zero_or_rounds_to_zero(
        self, px_per_meter, expected_dpi
    ):
        # -- exercise private helper directly; covers both legacy 0 case and
        #    the new rounds-to-zero guard --
        assert Bmp._dpi(px_per_meter) == expected_dpi

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def Bmp__init__(self, request):
        return initializer_mock(request, Bmp)
