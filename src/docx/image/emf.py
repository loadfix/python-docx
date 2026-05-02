"""EMF (Enhanced Metafile) image header parser.

Parses the ``EMR_HEADER`` record that begins every valid EMF stream. The
record layout is documented in [MS-EMF] section 2.3.4.2:

* offset 0x00  ``RecordType`` (uint32)   — always ``0x00000001`` for ``EMR_HEADER``.
* offset 0x04  ``RecordSize`` (uint32)
* offset 0x08  ``Bounds``     (RectL, 16 bytes) — picture frame in device units (pixels).
* offset 0x18  ``Frame``      (RectL, 16 bytes) — picture frame in 0.01-mm units.
* offset 0x28  ``Signature``  (uint32)   — ``0x464D4520`` (ASCII ``' EMF'``).
* offset 0x2C  ``Version``    (uint32)
* offset 0x30  ``Size``       (uint32)
* offset 0x34  ``Handles``    (uint16)
* offset 0x36  ``Reserved``   (uint16)
* offset 0x38  ``nDescription``, ``offDescription`` (uint32 each)
* offset 0x40  ``PalEntries``         (uint32)
* offset 0x44  ``Device``             (SizeL, 8 bytes) — reference device size in pixels.
* offset 0x4C  ``Millimeters``        (SizeL, 8 bytes) — reference device size in mm.
"""

from __future__ import annotations

from typing import IO

from docx.image.constants import MIME_TYPE
from docx.image.exceptions import InvalidImageStreamError
from docx.image.helpers import LITTLE_ENDIAN, StreamReader
from docx.image.image import BaseImageHeader


_EMR_HEADER_RECORD_TYPE = 0x00000001
_EMF_SIGNATURE = 0x464D4520  # " EMF"


class Emf(BaseImageHeader):
    """Image header parser for EMF (Enhanced Metafile) images.

    .. versionadded:: 1.3.0.dev0
    """

    @classmethod
    def from_stream(cls, stream: IO[bytes]) -> "Emf":
        """Return an |Emf| instance parsed from `stream`."""
        rdr = StreamReader(stream, LITTLE_ENDIAN)

        record_type = rdr.read_long(0x00)
        if record_type != _EMR_HEADER_RECORD_TYPE:
            raise InvalidImageStreamError(
                "not an EMF image (record type 0x%08X)" % record_type
            )
        signature = rdr.read_long(0x28)
        if signature != _EMF_SIGNATURE:
            raise InvalidImageStreamError(
                "invalid EMF signature 0x%08X" % signature
            )

        # -- Frame is a RectL in 0.01 mm units (HIMETRIC). Width and height
        #    derive from (right - left) and (bottom - top). --
        frame_left = _read_slong(rdr, 0x18)
        frame_top = _read_slong(rdr, 0x1C)
        frame_right = _read_slong(rdr, 0x20)
        frame_bottom = _read_slong(rdr, 0x24)

        frame_w_himetric = frame_right - frame_left
        frame_h_himetric = frame_bottom - frame_top

        # -- Reference device resolution: Device (pixels) / Millimeters. --
        device_w_px = rdr.read_long(0x44)
        device_h_px = rdr.read_long(0x48)
        device_w_mm = rdr.read_long(0x4C)
        device_h_mm = rdr.read_long(0x50)

        horz_dpi = _dpi(device_w_px, device_w_mm)
        vert_dpi = _dpi(device_h_px, device_h_mm)

        # -- Convert frame dimensions (0.01 mm) to pixels at 96 dpi so the
        #    values are useful when the device fields are absent/zero. --
        px_width = _himetric_to_px(frame_w_himetric)
        px_height = _himetric_to_px(frame_h_himetric)

        return cls(px_width, px_height, horz_dpi, vert_dpi)

    @property
    def content_type(self) -> str:
        """MIME content type for this image, always ``image/x-emf``."""
        return MIME_TYPE.EMF

    @property
    def default_ext(self) -> str:
        """Default filename extension, always ``'emf'``."""
        return "emf"


def _read_slong(rdr: StreamReader, offset: int) -> int:
    """Read a signed little-endian 32-bit integer at `offset`."""
    val = rdr.read_long(offset)
    if val & 0x80000000:
        val -= 0x1_0000_0000
    return val


def _dpi(px: int, mm: int) -> int:
    """Return integer dpi from a pixel count over a millimetre count."""
    if not px or not mm:
        return 96
    dpi = int(round(px * 25.4 / mm))
    return dpi if dpi > 0 else 96


def _himetric_to_px(himetric: int) -> int:
    """Convert HIMETRIC (0.01 mm) units to pixels at 96 dpi."""
    if himetric <= 0:
        return 0
    return int(round(himetric * 96 / 2540.0))
