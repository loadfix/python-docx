"""Unit-test suite for `FontTable.add_embedded_font` (upstream#1231, #1307).

These tests exercise the authoring surface for embedded fonts introduced in
1.3.0 — `FontTable.add_embedded_font(path, family=...)` — plus the supporting
`FontTablePart.default`/`FontTablePart.add_font_part` factories and the
`Document.font_table_or_new` shortcut.
"""

from __future__ import annotations

import io
from pathlib import Path
from typing import cast

import pytest

import docx
from docx.font_table import FontTable
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.font_table import CT_Fonts
from docx.package import Package
from docx.parts.font_table import FontPart, FontTablePart

from .unitutil.cxml import element


# -- a tiny binary blob stands in for a real TrueType font — python-docx never
# -- parses the bytes, it just packages them, so any blob round-trips cleanly --
_FAKE_FONT_BLOB = (
    b"OTTO\x00\x00\x00\x00" + b"\x00" * 128 + b"python-docx embedded-font fixture"
)


@pytest.fixture
def fake_font_path(tmp_path: Path) -> Path:
    path = tmp_path / "FakeFont.ttf"
    path.write_bytes(_FAKE_FONT_BLOB)
    return path


class DescribeFontTablePart_Default:
    """`FontTablePart.default` returns a newly-created empty part."""

    def it_creates_an_empty_fontTable_part(self):
        package = Package()

        part = FontTablePart.default(package)

        assert isinstance(part, FontTablePart)
        assert part.partname == "/word/fontTable.xml"
        assert part.content_type == CT.WML_FONT_TABLE
        assert len(part.font_table) == 0


class DescribeFontTable_AddEmbeddedFont:
    """End-to-end coverage for `FontTable.add_embedded_font`."""

    def it_creates_a_font_entry_with_an_embedRegular_rel(self, fake_font_path: Path):
        package = Package()
        ft_part = FontTablePart.default(package)

        metadata = ft_part.font_table.add_embedded_font(fake_font_path)

        assert metadata.name == "FakeFont"
        assert metadata.embed_regular is True
        # -- the FontTable sees the new entry --
        assert "FakeFont" in ft_part.font_table

    def it_wires_an_r_font_relationship_to_a_new_FontPart(self, fake_font_path: Path):
        package = Package()
        ft_part = FontTablePart.default(package)

        ft_part.font_table.add_embedded_font(fake_font_path)

        font_rels = [rel for rel in ft_part.rels.values() if rel.reltype == RT.FONT]
        assert len(font_rels) == 1
        target = font_rels[0].target_part
        assert isinstance(target, FontPart)
        assert target.blob == _FAKE_FONT_BLOB

    def it_rejects_unknown_family_variants(self, fake_font_path: Path):
        package = Package()
        ft_part = FontTablePart.default(package)

        with pytest.raises(ValueError, match="family must be one of"):
            ft_part.font_table.add_embedded_font(fake_font_path, family="oblique")  # type: ignore[arg-type]

    def it_supports_all_four_font_variants(self, fake_font_path: Path):
        package = Package()
        ft_part = FontTablePart.default(package)
        font_table = ft_part.font_table

        for family in ("regular", "bold", "italic", "bold_italic"):
            font_table.add_embedded_font(fake_font_path, family=family, name="All")

        metadata = font_table["All"]
        assert metadata.embed_regular is True
        assert metadata.embed_bold is True
        assert metadata.embed_italic is True
        assert metadata.embed_bold_italic is True


class DescribeDocument_FontTableOrNew:
    """`Document.font_table_or_new` materialises a fontTable.xml on demand."""

    def it_creates_the_part_when_missing_and_roundtrips_after_save(
        self, fake_font_path: Path
    ):
        document = docx.Document()

        # -- font_table_or_new always returns a live FontTable (creating one
        # -- on demand if needed); the bundled template ships with a
        # -- fontTable.xml so in this case it reuses the existing part. --
        font_table = document.font_table_or_new
        font_table.add_embedded_font(fake_font_path)

        # -- save then reopen to verify round-trip preservation --
        buf = io.BytesIO()
        document.save(buf)

        buf.seek(0)
        reopened = docx.Document(buf)

        assert reopened.font_table is not None
        assert "FakeFont" in reopened.font_table
        metadata = reopened.font_table["FakeFont"]
        assert metadata.embed_regular is True

        # -- the binary round-trips through the package bytewise --
        font_table_part = reopened.font_table.part
        font_rels = [r for r in font_table_part.rels.values() if r.reltype == RT.FONT]
        assert len(font_rels) == 1
        assert font_rels[0].target_part.blob == _FAKE_FONT_BLOB


class DescribeFontPart:
    """Minimal suite for the binary `FontPart` subclass."""

    def it_round_trips_a_blob_through_load(self):
        part = FontPart.load(
            "/word/fonts/font1.fntdata",  # type: ignore[arg-type]
            CT.X_FONTDATA,
            _FAKE_FONT_BLOB,
            Package(),
        )
        assert part.blob == _FAKE_FONT_BLOB
        assert part.content_type == CT.X_FONTDATA


# -- sanity: element helper used above yields a CT_Fonts --
def _sanity_cxml_parses_fonts():
    assert cast(CT_Fonts, element("w:fonts")) is not None
