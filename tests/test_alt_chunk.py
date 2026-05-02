# pyright: reportPrivateUsage=false

"""Unit test suite for the docx.alt_chunk module."""

from __future__ import annotations

from docx import Document
from docx.alt_chunk import AltChunk
from docx.document import Document as DocumentCls
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.parts.alt_chunk import AltChunkPart, _ext_for_content_type


class DescribeDocumentAddAltChunk:
    """Unit-test suite for `Document.add_alt_chunk`."""

    def it_returns_an_AltChunk_proxy(self):
        document: DocumentCls = Document()

        alt_chunk = document.add_alt_chunk("<p>hello</p>")

        assert isinstance(alt_chunk, AltChunk)

    def it_appends_a_w_altChunk_element_to_the_body(self):
        document: DocumentCls = Document()

        document.add_alt_chunk(b"<p>hi</p>")

        body = document._element.body
        assert len(body.altChunk_lst) == 1

    def it_creates_a_relationship_with_aFChunk_reltype(self):
        document: DocumentCls = Document()

        alt_chunk = document.add_alt_chunk("<p>hi</p>")

        # -- the rId on the altChunk element resolves to an AltChunkPart --
        assert alt_chunk.rId is not None
        assert isinstance(alt_chunk.part, AltChunkPart)
        # -- and the relationship type is aFChunk --
        document_part = document._part
        assert document_part.rels[alt_chunk.rId].reltype == RT.A_F_CHUNK

    def it_encodes_str_content_as_utf_8_bytes(self):
        document: DocumentCls = Document()

        alt_chunk = document.add_alt_chunk("café", content_type="text/plain")

        assert alt_chunk.content == "café".encode("utf-8")

    def it_defaults_the_content_type_to_text_html(self):
        document: DocumentCls = Document()

        alt_chunk = document.add_alt_chunk("<p>hi</p>")

        assert alt_chunk.content_type == "text/html"

    def it_accepts_a_custom_content_type(self):
        document: DocumentCls = Document()

        alt_chunk = document.add_alt_chunk(b"{\\rtf1}", content_type="application/rtf")

        assert alt_chunk.content_type == "application/rtf"


class DescribeDocumentAltChunks:
    """Unit-test suite for `Document.alt_chunks`."""

    def it_returns_an_empty_list_when_there_are_no_altChunks(self):
        document: DocumentCls = Document()

        assert document.alt_chunks == []

    def it_lists_one_proxy_per_altChunk_in_document_order(self):
        document: DocumentCls = Document()
        document.add_alt_chunk("<p>first</p>")
        document.add_alt_chunk("<p>second</p>", content_type="text/html")

        chunks = document.alt_chunks

        assert len(chunks) == 2
        assert all(isinstance(ch, AltChunk) for ch in chunks)
        assert chunks[0].content == b"<p>first</p>"
        assert chunks[1].content == b"<p>second</p>"

    def it_round_trips_through_save_and_open(self, tmp_path):
        document: DocumentCls = Document()
        document.add_alt_chunk("<p>hello</p>", content_type="text/html")
        path = tmp_path / "roundtrip.docx"
        document.save(str(path))

        reopened: DocumentCls = Document(str(path))

        chunks = reopened.alt_chunks
        assert len(chunks) == 1
        assert chunks[0].content_type == "text/html"
        assert chunks[0].content == b"<p>hello</p>"


class DescribeAltChunkPart:
    """Unit-test suite for `docx.parts.alt_chunk.AltChunkPart`."""

    def it_picks_a_partname_extension_from_the_content_type(self):
        assert _ext_for_content_type("text/html") == ".html"
        assert _ext_for_content_type("application/rtf") == ".rtf"
        assert _ext_for_content_type("text/rtf") == ".rtf"
        assert _ext_for_content_type("application/xhtml+xml") == ".xhtml"
        assert _ext_for_content_type("text/plain") == ".txt"
        assert _ext_for_content_type("application/msword") == ".doc"
        assert _ext_for_content_type("weird/thing") == ".bin"

    def it_can_be_loaded_from_blob(self):
        # -- simulate the PartFactory.load path --
        from docx.opc.packuri import PackURI

        part = AltChunkPart.load(
            PackURI("/word/afchunk1.html"),
            "text/html",
            b"<p>x</p>",
            None,  # type: ignore[arg-type]
        )
        assert part.blob == b"<p>x</p>"
        assert part.content_type == "text/html"
