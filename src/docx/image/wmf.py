"""WMF (Windows Metafile) image header parser.

Only the **Aldus placeable** WMF variant carries embedded dimensions, so
non-placeable WMF streams (``0x9AC6CDD7`` magic missing) are rejected.

Placeable header layout (22 bytes) per MS-WMF 2.3.2.3:

* offset 0x00  ``Key``       uint32  ``0x9AC6CDD7``
* offset 0x04  ``HWmf``      uint16  (unused, 0)
* offset 0x06  ``BoundingBox`` 4 x int16  — left, top, right, bottom in TWIPS units
  scaled by ``Inch``.
* offset 0x0E  ``Inch``      uint16  — number of ``BoundingBox`` units per inch.
* offset 0x10  ``Reserved``  uint32
* offset 0x14  ``Checksum``  uint16
"""

from __future__ import annotations

from typing import IO

from docx.image.constants import MIME_TYPE
from docx.image.exceptions import InvalidImageStreamError
from docx.image.helpers import LITTLE_ENDIAN, StreamReader
from docx.image.image import BaseImageHeader


WMF_PLACEABLE_MAGIC = 0x9AC6CDD7


class Wmf(BaseImageHeader):
    """Image header parser for placeable WMF (Windows Metafile) images.

    .. versionadded:: 2026.05.0
    """

    @classmethod
    def from_stream(cls, stream: IO[bytes]) -> "Wmf":
        """Return a |Wmf| instance parsed from `stream`."""
        rdr = StreamReader(stream, LITTLE_ENDIAN)

        key = rdr.read_long(0x00)
        if key != WMF_PLACEABLE_MAGIC:
            raise InvalidImageStreamError(
                "WMF stream is not placeable (missing 0x9AC6CDD7 magic); "
                "dimensions cannot be determined"
            )

        left = _read_sshort(rdr, 0x06)
        top = _read_sshort(rdr, 0x08)
        right = _read_sshort(rdr, 0x0A)
        bottom = _read_sshort(rdr, 0x0C)
        inch = rdr.read_short(0x0E)

        if not inch:
            raise InvalidImageStreamError(
                "WMF placeable header has zero Inch field"
            )

        width_units = right - left
        height_units = bottom - top
        # -- bounding-box units are 1/Inch per unit; px at 96 dpi --
        px_width = int(round(abs(width_units) * 96 / inch)) if width_units else 0
        px_height = int(round(abs(height_units) * 96 / inch)) if height_units else 0
        horz_dpi = inch
        vert_dpi = inch

        return cls(px_width, px_height, horz_dpi, vert_dpi)

    @property
    def content_type(self) -> str:
        """MIME content type for this image, always ``image/x-wmf``."""
        return MIME_TYPE.WMF

    @property
    def default_ext(self) -> str:
        """Default filename extension, always ``'wmf'``."""
        return "wmf"


def _read_sshort(rdr: StreamReader, offset: int) -> int:
    """Read a signed little-endian 16-bit integer at `offset`."""
    val = rdr.read_short(offset)
    if val & 0x8000:
        val -= 0x1_0000
    return val
