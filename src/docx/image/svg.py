"""SVG image header parser."""

from __future__ import annotations

import re
import struct
import zlib
from typing import IO

from docx.image.constants import MIME_TYPE
from docx.image.image import BaseImageHeader


class Svg(BaseImageHeader):
    """Image header parser for SVG images."""

    @property
    def content_type(self) -> str:
        return MIME_TYPE.SVG

    @property
    def default_ext(self) -> str:
        return "svg"

    @classmethod
    def from_stream(cls, stream: IO[bytes]) -> Svg:
        stream.seek(0)
        data = stream.read()
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError:
            return cls(300, 150, 96, 96)
        px_width, px_height = cls._parse_dimensions(text)
        # SVG uses 96 DPI (CSS reference pixel)
        return cls(px_width, px_height, 96, 96)

    @classmethod
    def _parse_dimensions(cls, svg_text: str) -> tuple[int, int]:
        import defusedxml.ElementTree as SafeET

        try:
            root = SafeET.fromstring(svg_text)
        except Exception:
            return 300, 150  # default SVG dimensions per spec

        # Check for width/height attributes
        width_str = root.get("width", "")
        height_str = root.get("height", "")

        width = cls._parse_length(width_str)
        height = cls._parse_length(height_str)

        if width and height:
            return width, height

        # Fall back to viewBox
        viewbox = root.get("viewBox", "")
        if viewbox:
            parts = re.split(r"[\s,]+", viewbox.strip())
            if len(parts) == 4:
                try:
                    vb_width = float(parts[2])
                    vb_height = float(parts[3])
                    if vb_width > 0 and vb_height > 0:
                        return int(round(vb_width)), int(round(vb_height))
                except ValueError:
                    pass

        return 300, 150

    @classmethod
    def _parse_length(cls, length_str: str | None) -> int | None:
        if not length_str:
            return None

        match = re.match(r"^\s*([\d.]+)\s*(px|pt|in|cm|mm|)\s*$", length_str)
        if not match:
            return None

        value = float(match.group(1))
        unit = match.group(2)

        if unit in ("", "px"):
            return int(round(value))
        elif unit == "pt":
            return int(round(value * 96 / 72))
        elif unit == "in":
            return int(round(value * 96))
        elif unit == "cm":
            return int(round(value * 96 / 2.54))
        elif unit == "mm":
            return int(round(value * 96 / 25.4))


def is_svg_stream(stream: IO[bytes]) -> bool:
    """Return True if `stream` contains an SVG image."""
    stream.seek(0)
    header = stream.read(4096)
    stripped = header.lstrip()
    # Strip BOM if present
    if stripped.startswith(b"\xef\xbb\xbf"):
        stripped = stripped[3:].lstrip()
    return stripped.startswith(b"<svg") or (
        stripped.startswith(b"<?xml") and b"<svg" in stripped
    )


def generate_fallback_png() -> bytes:
    """Generate a minimal 1x1 transparent PNG for SVG fallback."""
    width, height = 1, 1
    ihdr_data = struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0)
    raw = b"\x00" + b"\x00\x00\x00\x00"
    compressed = zlib.compress(raw)

    def chunk(type_code: bytes, data: bytes) -> bytes:
        c = type_code + data
        crc = struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)
        return struct.pack(">I", len(data)) + c + crc

    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr_data)
        + chunk(b"IDAT", compressed)
        + chunk(b"IEND", b"")
    )
