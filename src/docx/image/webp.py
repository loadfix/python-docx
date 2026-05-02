"""WebP image header parser.

Supports the three VP8 chunk variants defined by the WebP specification:

* ``VP8 `` — lossy (simple file format).
* ``VP8L`` — lossless.
* ``VP8X`` — extended file format, used for animation / alpha / etc.

See https://developers.google.com/speed/webp/docs/riff_container for the
authoritative container layout used below.
"""

from __future__ import annotations

from typing import IO

from docx.image.constants import MIME_TYPE
from docx.image.exceptions import InvalidImageStreamError
from docx.image.helpers import LITTLE_ENDIAN, StreamReader
from docx.image.image import BaseImageHeader


class WebP(BaseImageHeader):
    """Image header parser for WebP images.

    The WebP format does not carry resolution (DPI) metadata, so both
    horizontal and vertical DPI default to 72 (matching GIF behaviour).

    .. versionadded:: 1.3.0.dev0
    """

    @classmethod
    def from_stream(cls, stream: IO[bytes]) -> "WebP":
        """Return a |WebP| instance parsed from WebP image in `stream`."""
        px_width, px_height = _WebpParser.parse_dimensions(stream)
        return cls(px_width, px_height, 72, 72)

    @property
    def content_type(self) -> str:
        """MIME content type for this image, always ``image/webp``."""
        return MIME_TYPE.WEBP

    @property
    def default_ext(self) -> str:
        """Default filename extension, always ``'webp'``."""
        return "webp"


class _WebpParser:
    """Extracts canvas dimensions from the first VP8/VP8L/VP8X chunk."""

    @classmethod
    def parse_dimensions(cls, stream: IO[bytes]) -> tuple[int, int]:
        rdr = StreamReader(stream, LITTLE_ENDIAN)

        # -- RIFF header ---------------------------------------------------
        riff_tag = rdr.read_str(4, 0)
        # (file size long at offset 4 is ignored)
        webp_tag = rdr.read_str(4, 8)
        if riff_tag != "RIFF" or webp_tag != "WEBP":
            raise InvalidImageStreamError("not a WebP image (missing RIFF/WEBP)")

        # -- First sub-chunk at offset 12 ----------------------------------
        chunk_fourcc = rdr.read_str(4, 12)
        # chunk payload size long at offset 16 is read but not required
        if chunk_fourcc == "VP8 ":
            return cls._parse_vp8(rdr)
        if chunk_fourcc == "VP8L":
            return cls._parse_vp8l(rdr)
        if chunk_fourcc == "VP8X":
            return cls._parse_vp8x(rdr)
        raise InvalidImageStreamError(
            "unrecognized WebP chunk type: %r" % chunk_fourcc
        )

    @staticmethod
    def _parse_vp8(rdr: StreamReader) -> tuple[int, int]:
        """Lossy VP8 bitstream. Width and height are 14-bit values located at
        byte offsets 26 and 28 respectively, after a 3-byte frame tag and the
        three start-code bytes ``0x9d 0x01 0x2a``."""
        # Sanity-check the start code (offset 23..25) to avoid mis-parsing
        # a malformed stream.
        start0 = rdr.read_byte(23)
        start1 = rdr.read_byte(24)
        start2 = rdr.read_byte(25)
        if (start0, start1, start2) != (0x9D, 0x01, 0x2A):
            raise InvalidImageStreamError("invalid VP8 start code")
        px_width = rdr.read_short(26) & 0x3FFF
        px_height = rdr.read_short(28) & 0x3FFF
        return px_width, px_height

    @staticmethod
    def _parse_vp8l(rdr: StreamReader) -> tuple[int, int]:
        """Lossless VP8L bitstream. After the 1-byte signature ``0x2f`` at
        offset 20, width-1 and height-1 are packed into a little-endian
        32-bit value as 14+14+1+3 bits (width, height, alpha_hint, version).
        """
        signature = rdr.read_byte(20)
        if signature != 0x2F:
            raise InvalidImageStreamError("invalid VP8L signature")
        bits = rdr.read_long(21)
        px_width = (bits & 0x3FFF) + 1
        px_height = ((bits >> 14) & 0x3FFF) + 1
        return px_width, px_height

    @staticmethod
    def _parse_vp8x(rdr: StreamReader) -> tuple[int, int]:
        """Extended VP8X chunk. Canvas width-1 and height-1 are stored as
        little-endian 24-bit unsigned integers at offsets 24 and 27."""
        # Read three bytes and assemble as little-endian.
        w0 = rdr.read_byte(24)
        w1 = rdr.read_byte(25)
        w2 = rdr.read_byte(26)
        h0 = rdr.read_byte(27)
        h1 = rdr.read_byte(28)
        h2 = rdr.read_byte(29)
        px_width = (w0 | (w1 << 8) | (w2 << 16)) + 1
        px_height = (h0 | (h1 << 8) | (h2 << 16)) + 1
        return px_width, px_height
