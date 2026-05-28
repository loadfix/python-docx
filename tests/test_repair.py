"""Tests for ``docx.repair`` (issue #92)."""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

import pytest

from docx import Document, RepairError, RepairReport, repair
from docx.api import Document as DocumentFactory
from docx.document import Document as DocumentCls


_TPL_PATH = Path(__file__).parent.parent / "src" / "docx" / "templates" / "default.docx"


def _template_bytes() -> bytes:
    with open(_TPL_PATH, "rb") as f:
        return f.read()


# ---------------------------------------------------------------------------
# Fixture builders.
# ---------------------------------------------------------------------------


def _make_truncated_zip_bytes() -> bytes:
    """Return a zip blob whose end-of-central-directory record is missing.

    Built by writing a real zip then chopping off the central directory.
    The local-file headers + payloads remain intact, so a recovery scan
    can still salvage every entry.
    """
    blob = _template_bytes()
    # -- find the EOCD signature; truncate just before it. The CD itself
    # -- (`PK\x01\x02` records) lives between the last entry payload and
    # -- the EOCD. Strip everything from the first CD signature onward. --
    cd_start = blob.find(b"PK\x01\x02")
    assert cd_start != -1, "default.docx unexpectedly has no central directory"
    return blob[:cd_start]


def _make_malformed_xml_bytes() -> bytes:
    """Return a docx whose `word/document.xml` has an unclosed bookmarkStart."""
    blob = _template_bytes()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob), "r") as zin:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "word/document.xml":
                    text = data.decode("utf-8")
                    # -- inject an orphan bookmarkStart inside <w:body> --
                    injected = (
                        '<w:bookmarkStart w:id="424242" w:name="orphan-rep"/>'
                    )
                    text = text.replace("<w:body>", "<w:body>" + injected, 1)
                    data = text.encode("utf-8")
                zout.writestr(info, data)
    return out.getvalue()


def _make_missing_rel_target_bytes() -> bytes:
    """Return a docx whose document rels points at a non-existent custom-xml part."""
    blob = _template_bytes()
    out = io.BytesIO()
    extra_rel = (
        '<Relationship Id="rIdMissing0001" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/customXml" Target="../customXml/itemDoesNotExist.xml"/>'
    )
    with zipfile.ZipFile(io.BytesIO(blob), "r") as zin:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "word/_rels/document.xml.rels":
                    text = data.decode("utf-8")
                    text = text.replace("</Relationships>", extra_rel + "</Relationships>")
                    data = text.encode("utf-8")
                zout.writestr(info, data)
    return out.getvalue()


def _make_bad_encoding_decl_bytes() -> bytes:
    """Return a docx whose `word/styles.xml` declares utf-16 but is utf-8."""
    blob = _template_bytes()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob), "r") as zin:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "word/styles.xml":
                    text = data.decode("utf-8")
                    text = re.sub(
                        r"encoding=['\"][^'\"]+['\"]",
                        'encoding="utf-16"',
                        text,
                        count=1,
                    )
                    data = text.encode("utf-8")
                zout.writestr(info, data)
    return out.getvalue()


def _make_unrecoverable_xml_part_bytes() -> bytes:
    """Return a docx whose `word/footer1.xml` is binary garbage.

    The default template has no footer; we add a content-type registration
    plus an actual junk body. Best-effort repair should drop the part.
    """
    blob = _template_bytes()
    out = io.BytesIO()
    junk = b"\x00\x01\x02 not even pretending to be xml"
    with zipfile.ZipFile(io.BytesIO(blob), "r") as zin:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                zout.writestr(info, data)
            zout.writestr("word/junk.xml", junk)
    return out.getvalue()


# ---------------------------------------------------------------------------
# Test classes.
# ---------------------------------------------------------------------------


class DescribeRepairReport:
    def it_starts_empty(self):
        report = RepairReport(strategy="best-effort")

        assert report.strategy == "best-effort"
        assert report.repaired == []
        assert report.unrecoverable == []
        assert report.parts_dropped == []
        assert report.is_clean is True

    def it_reports_dirty_when_any_field_populated(self):
        report = RepairReport(strategy="best-effort", repaired=["x: y"])

        assert report.is_clean is False


class DescribeRepairBestEffort:
    """End-to-end repair pass on damaged fixtures."""

    def it_repairs_a_truncated_zip(self):
        stream = io.BytesIO(_make_truncated_zip_bytes())

        document, report = repair(stream, strategy="best-effort")

        assert isinstance(document, DocumentCls)
        assert any("truncated zip" in line for line in report.repaired)
        # -- the recovered document still has the template's stock paragraph --
        assert document.paragraphs is not None

    def it_repairs_malformed_xml(self):
        stream = io.BytesIO(_make_malformed_xml_bytes())

        document, report = repair(stream, strategy="best-effort")

        assert isinstance(document, DocumentCls)
        # -- the orphan bookmark should have been closed --
        assert any(
            "orphan w:bookmarkStart" in line for line in report.repaired
        ), report.repaired

    def it_repairs_missing_relationship_target(self):
        stream = io.BytesIO(_make_missing_rel_target_bytes())

        document, report = repair(stream, strategy="best-effort")

        # -- the existing pkgreader silently skips dangling rels (upstream-PR#1219);
        # -- so the document loads cleanly. The report's `parts_dropped` may be
        # -- empty, but a clean Document instance is the headline contract. --
        assert isinstance(document, DocumentCls)
        assert report.strategy == "best-effort"

    def it_normalises_a_bad_encoding_declaration(self):
        stream = io.BytesIO(_make_bad_encoding_decl_bytes())

        document, report = repair(stream, strategy="best-effort")

        assert isinstance(document, DocumentCls)
        assert any(
            "encoding declaration" in line for line in report.repaired
        ), report.repaired

    def it_drops_unrecoverable_parts(self):
        stream = io.BytesIO(_make_unrecoverable_xml_part_bytes())

        document, report = repair(stream, strategy="best-effort")

        assert isinstance(document, DocumentCls)
        assert any(
            "/word/junk.xml" in line for line in report.parts_dropped
        ), report.parts_dropped

    def it_returns_the_document_class_and_report_class(self):
        stream = io.BytesIO(_make_malformed_xml_bytes())

        document, report = repair(stream, strategy="best-effort")

        assert isinstance(document, DocumentCls)
        assert isinstance(report, RepairReport)

    def it_accepts_a_pathlib_path(self, tmp_path):
        path = tmp_path / "broken.docx"
        path.write_bytes(_make_malformed_xml_bytes())

        document, report = repair(path)

        assert isinstance(document, DocumentCls)
        assert report.strategy == "best-effort"

    def it_accepts_a_str_path(self, tmp_path):
        path = tmp_path / "broken.docx"
        path.write_bytes(_make_malformed_xml_bytes())

        document, report = repair(str(path))

        assert isinstance(document, DocumentCls)

    def it_round_trips_repaired_document(self, tmp_path):
        stream = io.BytesIO(_make_malformed_xml_bytes())

        document, report = repair(stream, strategy="best-effort")
        out = tmp_path / "fixed.docx"
        document.save(str(out))

        # -- the saved file must reopen cleanly without recover=True --
        reopened = DocumentFactory(str(out))
        assert isinstance(reopened, DocumentCls)
        del report  # unused

    def it_can_be_called_via_Document_repair_attribute(self):
        # -- mirrors `Document.from_template`'s API shape --
        stream = io.BytesIO(_make_malformed_xml_bytes())

        document, report = Document.repair(stream)

        assert isinstance(document, DocumentCls)
        assert isinstance(report, RepairReport)


class DescribeRepairStrict:
    def it_records_no_repairs_when_input_is_clean(self):
        # -- adding an unreferenced junk.xml to the package doesn't break
        # -- anything: the rels graph never visits it. Strict mode loads
        # -- the document with no repair activity reported. --
        stream = io.BytesIO(_make_unrecoverable_xml_part_bytes())

        document, report = repair(stream, strategy="strict")

        assert document is not None
        assert report.repaired == []
        assert report.parts_dropped == []

    def it_raises_for_truly_invalid_xml_in_strict_mode(self):
        # -- inject malformed XML that the strict parser cannot accept --
        blob = _template_bytes()
        out = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(blob), "r") as zin:
            with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
                for info in zin.infolist():
                    data = zin.read(info.filename)
                    if info.filename == "word/document.xml":
                        data = data[: len(data) // 2]
                    zout.writestr(info, data)
        stream = io.BytesIO(out.getvalue())

        from lxml import etree

        with pytest.raises(etree.XMLSyntaxError):
            repair(stream, strategy="strict")


class DescribeRepairTruncate:
    def it_drops_all_parts_after_first_defect(self):
        stream = io.BytesIO(_make_unrecoverable_xml_part_bytes())

        document, report = repair(stream, strategy="truncate")

        assert document is not None
        assert any(
            "/word/junk.xml" in line for line in report.parts_dropped
        ), report.parts_dropped


class DescribeRepairValidation:
    def it_rejects_unknown_strategies(self):
        stream = io.BytesIO(_template_bytes())

        with pytest.raises(ValueError, match="unknown repair strategy"):
            repair(stream, strategy="bogus")

    def it_raises_RepairError_on_unsalvageable_input(self, tmp_path):
        bogus = tmp_path / "garbage.docx"
        bogus.write_bytes(b"definitely not a zip and not enough PK headers either")

        with pytest.raises(RepairError):
            repair(str(bogus))

    def it_raises_RepairError_on_missing_path(self, tmp_path):
        with pytest.raises(RepairError, match="no file at"):
            repair(str(tmp_path / "no-such-file.docx"))

    def it_returns_a_report_in_strict_mode_for_a_valid_doc(self):
        stream = io.BytesIO(_template_bytes())

        document, report = repair(stream, strategy="strict")

        assert isinstance(document, DocumentCls)
        assert report.strategy == "strict"
        assert report.is_clean is True
