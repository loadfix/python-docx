"""Unit test suite for docx.image.svg module."""

import io

import pytest

from docx.image.constants import MIME_TYPE
from docx.image.image import _ImageHeaderFactory
from docx.image.svg import Svg, generate_fallback_png, is_svg_stream


class DescribeSvg:
    def it_can_construct_from_a_stream(self):
        svg_bytes = (
            b'<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100">'
            b"</svg>"
        )
        stream = io.BytesIO(svg_bytes)
        svg = Svg.from_stream(stream)
        assert svg.px_width == 200
        assert svg.px_height == 100
        assert svg.content_type == MIME_TYPE.SVG
        assert svg.default_ext == "svg"

    def it_parses_dimensions_from_width_and_height_attrs(self):
        svg_bytes = (
            b'<svg xmlns="http://www.w3.org/2000/svg" width="300" height="200">'
            b"</svg>"
        )
        stream = io.BytesIO(svg_bytes)
        svg = Svg.from_stream(stream)
        assert svg.px_width == 300
        assert svg.px_height == 200

    def it_parses_dimensions_from_viewBox_when_no_width_height(self):
        svg_bytes = (
            b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300">'
            b"</svg>"
        )
        stream = io.BytesIO(svg_bytes)
        svg = Svg.from_stream(stream)
        assert svg.px_width == 400
        assert svg.px_height == 300

    def it_parses_width_and_height_with_units(self):
        svg_bytes = (
            b'<svg xmlns="http://www.w3.org/2000/svg" width="2in" height="1in">'
            b"</svg>"
        )
        stream = io.BytesIO(svg_bytes)
        svg = Svg.from_stream(stream)
        assert svg.px_width == 192  # 2 * 96
        assert svg.px_height == 96   # 1 * 96

    def it_uses_default_dimensions_when_no_size_info(self):
        svg_bytes = (
            b'<svg xmlns="http://www.w3.org/2000/svg">'
            b"</svg>"
        )
        stream = io.BytesIO(svg_bytes)
        svg = Svg.from_stream(stream)
        assert svg.px_width == 300
        assert svg.px_height == 150

    def it_uses_96_dpi(self):
        svg_bytes = (
            b'<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">'
            b"</svg>"
        )
        stream = io.BytesIO(svg_bytes)
        svg = Svg.from_stream(stream)
        assert svg.horz_dpi == 96
        assert svg.vert_dpi == 96


class Describe_is_svg_stream:
    def it_returns_True_for_an_svg_stream(self):
        svg_bytes = b'<svg xmlns="http://www.w3.org/2000/svg"></svg>'
        stream = io.BytesIO(svg_bytes)
        assert is_svg_stream(stream) is True

    def it_returns_True_for_svg_with_xml_declaration(self):
        svg_bytes = (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<svg xmlns="http://www.w3.org/2000/svg"></svg>'
        )
        stream = io.BytesIO(svg_bytes)
        assert is_svg_stream(stream) is True

    def it_returns_True_for_svg_with_BOM(self):
        svg_bytes = (
            b"\xef\xbb\xbf"
            b'<svg xmlns="http://www.w3.org/2000/svg"></svg>'
        )
        stream = io.BytesIO(svg_bytes)
        assert is_svg_stream(stream) is True

    def it_returns_False_for_a_non_svg_stream(self):
        stream = io.BytesIO(b"not an svg file at all")
        assert is_svg_stream(stream) is False

    def it_returns_False_for_non_svg_xml(self):
        stream = io.BytesIO(b'<?xml version="1.0"?><html></html>')
        assert is_svg_stream(stream) is False


class Describe_generate_fallback_png:
    def it_generates_a_valid_png(self):
        png_bytes = generate_fallback_png()
        assert png_bytes[:8] == b"\x89PNG\r\n\x1a\n"
        assert len(png_bytes) > 8


class Describe_ImageHeaderFactory_SVG:
    def it_returns_Svg_for_an_svg_stream(self):
        svg_bytes = (
            b'<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100">'
            b"</svg>"
        )
        stream = io.BytesIO(svg_bytes)
        image_header = _ImageHeaderFactory(stream)
        assert isinstance(image_header, Svg)

    def it_returns_Svg_for_svg_with_xml_declaration(self):
        svg_bytes = (
            b'<?xml version="1.0" encoding="UTF-8"?>\n'
            b'<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100">'
            b"</svg>"
        )
        stream = io.BytesIO(svg_bytes)
        image_header = _ImageHeaderFactory(stream)
        assert isinstance(image_header, Svg)
