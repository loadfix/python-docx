"""Provides objects that can characterize image streams.

That characterization is as to content type and size, as a required step in including
them in a document.
"""

from docx.image.bmp import Bmp
from docx.image.emf import Emf
from docx.image.gif import Gif
from docx.image.jpeg import Exif, Jfif
from docx.image.png import Png
from docx.image.tiff import Tiff
from docx.image.webp import WebP
from docx.image.wmf import Wmf

SIGNATURES = (
    # class, offset, signature_bytes
    (Png, 0, b"\x89PNG\x0d\x0a\x1a\x0a"),
    (Jfif, 6, b"JFIF"),
    (Exif, 6, b"Exif"),
    (Gif, 0, b"GIF87a"),
    (Gif, 0, b"GIF89a"),
    (Tiff, 0, b"MM\x00*"),  # big-endian (Motorola) TIFF
    (Tiff, 0, b"II*\x00"),  # little-endian (Intel) TIFF
    (Bmp, 0, b"BM"),
    # -- EMF: EMR_HEADER record type 0x00000001 plus the ASCII ' EMF'
    #    signature at 0x28 uniquely identify an EMF stream. --
    (Emf, 0x28, b" EMF"),
    # -- Placeable WMF magic (little-endian 0x9AC6CDD7). Plain WMF
    #    streams (01 00 09 00 / 02 00 09 00) are rejected because they
    #    do not carry dimensions. --
    (Wmf, 0, b"\xd7\xcd\xc6\x9a"),
)

# -- Classes referenced by the detection helper in ``docx.image.image``
#    but not covered by the fixed-offset SIGNATURES table above. ---------
__all__ = ("SIGNATURES", "WebP", "Emf", "Wmf")
