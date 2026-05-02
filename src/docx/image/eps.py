"""EPS (Encapsulated PostScript) image header parser.

EPS is a text-based format. We accept two on-disk shapes:

1. Plain EPS beginning with ``%!PS-Adobe``.
2. DOS EPS binary ("EPSF") beginning with the 4-byte magic
   ``C5 D0 D3 C6``; the PostScript section start/length is read from the
   binary header and dimensions are extracted from the ``%%BoundingBox``
   comment within it.

Dimensions are read from the ``%%BoundingBox: llx lly urx ury`` DSC
comment; values are in PostScript points (1/72 inch). We therefore report
``horz_dpi``/``vert_dpi`` = 72 and convert the point dimensions to pixels
at that dpi so ``px_width / horz_dpi`` yields the native size in inches.

.. versionadded:: 2026.05.0
"""

from __future__ import annotations

import re
import struct
from typing import IO

from docx.image.constants import MIME_TYPE
from docx.image.exceptions import InvalidImageStreamError
from docx.image.image import BaseImageHeader


_DOS_EPS_MAGIC = b"\xc5\xd0\xd3\xc6"
_BOUNDING_BOX_RE = re.compile(
    rb"%%BoundingBox:\s*(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)"
)


class Eps(BaseImageHeader):
    """Image header parser for EPS (Encapsulated PostScript) images.

    .. versionadded:: 2026.05.0
    """

    @classmethod
    def from_stream(cls, stream: IO[bytes]) -> "Eps":
        """Return an |Eps| instance parsed from `stream`."""
        stream.seek(0)
        header = stream.read(64 * 1024)
        ps_text = _extract_postscript(header)
        if not ps_text.startswith(b"%!PS-Adobe"):
            raise InvalidImageStreamError(
                "not an EPS stream (missing %!PS-Adobe header)"
            )
        match = _BOUNDING_BOX_RE.search(ps_text)
        if match is None:
            raise InvalidImageStreamError(
                "EPS stream has no %%BoundingBox comment"
            )
        llx, lly, urx, ury = (float(v) for v in match.groups())
        width_pts = urx - llx
        height_pts = ury - lly
        px_width = max(int(round(width_pts)), 0)
        px_height = max(int(round(height_pts)), 0)
        return cls(px_width, px_height, 72, 72)

    @property
    def content_type(self) -> str:
        """MIME content type for this image, always ``application/postscript``."""
        return MIME_TYPE.EPS

    @property
    def default_ext(self) -> str:
        """Default filename extension, always ``'eps'``."""
        return "eps"


def _extract_postscript(header: bytes) -> bytes:
    """Return the PostScript payload from `header`.

    Handles both plain EPS and DOS EPS ("EPSF") containers. Falls through
    to the raw bytes when no DOS-EPS magic is present.
    """
    if header.startswith(_DOS_EPS_MAGIC) and len(header) >= 12:
        ps_offset, ps_length = struct.unpack_from("<II", header, 4)
        end = ps_offset + ps_length
        return header[ps_offset:end]
    return header


def is_eps_stream(stream: IO[bytes]) -> bool:
    """Return True if `stream` contains an EPS image."""
    stream.seek(0)
    header = stream.read(4096)
    if header.startswith(_DOS_EPS_MAGIC):
        return True
    stripped = header.lstrip()
    return stripped.startswith(b"%!PS-Adobe")
