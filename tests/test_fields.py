"""Unit-test suite for the docx.fields module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.fields import Field, WD_FIELD_TYPE
from docx.oxml.fields import CT_FldSimple
from docx.oxml.ns import qn
from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.run import CT_R

from .unitutil.cxml import element


class DescribeWD_FIELD_TYPE:
    """Sanity check for the constant set."""

    @pytest.mark.parametrize(
        ("name", "value"),
        [
            ("PAGE", "PAGE"),
            ("NUMPAGES", "NUMPAGES"),
            ("DATE", "DATE"),
            ("TIME", "TIME"),
            ("AUTHOR", "AUTHOR"),
            ("REF", "REF"),
            ("TOC", "TOC"),
            ("SEQ", "SEQ"),
            ("HYPERLINK", "HYPERLINK"),
            ("PAGEREF", "PAGEREF"),
        ],
    )
    def it_exposes_common_field_types_as_string_constants(self, name: str, value: str):
        assert getattr(WD_FIELD_TYPE, name) == value


class DescribeField_Simple:
    """Unit-test suite for `Field` wrapping a ``w:fldSimple`` element."""

    def it_is_not_complex(self):
        fldSimple = cast(CT_FldSimple, element('w:fldSimple{w:instr=PAGE}'))
        field = Field.for_simple(fldSimple)
        assert field.is_complex is False

    def it_exposes_the_raw_instruction(self):
        # -- backslashes aren't supported by the cxml attribute grammar; build
        #    the element via OxmlElement and set the w:instr attribute directly.
        from docx.oxml.ns import qn
        from docx.oxml.parser import OxmlElement

        fldSimple = cast(CT_FldSimple, OxmlElement("w:fldSimple"))
        fldSimple.set(qn("w:instr"), "REF bookmark1 \\h")
        field = Field.for_simple(fldSimple)
        assert field.instruction == "REF bookmark1 \\h"

    def it_returns_empty_instruction_when_attr_missing(self):
        # -- w:instr is required, but defensively handle absence --
        from docx.oxml.parser import OxmlElement

        fldSimple = cast(CT_FldSimple, OxmlElement("w:fldSimple"))
        field = Field.for_simple(fldSimple)
        assert field.instruction == ""

    @pytest.mark.parametrize(
        ("instr", "expected_type"),
        [
            ("PAGE", "PAGE"),
            ("NUMPAGES", "NUMPAGES"),
            ("REF bookmark1 \\h", "REF"),
            ("TOC \\o \"1-3\"", "TOC"),
            ("SEQ Table", "SEQ"),
            ("DATE", "DATE"),
            ("TIME \\@ \"h:mm AM/PM\"", "TIME"),
            ("AUTHOR", "AUTHOR"),
            ("HYPERLINK \"https://example.com\"", "HYPERLINK"),
            ("PAGEREF _Toc12345", "PAGEREF"),
        ],
    )
    def it_parses_the_type_from_the_instruction(self, instr: str, expected_type: str):
        from docx.oxml.parser import OxmlElement
        from docx.oxml.ns import qn

        fldSimple = cast(CT_FldSimple, OxmlElement("w:fldSimple"))
        fldSimple.set(qn("w:instr"), instr)
        field = Field.for_simple(fldSimple)
        assert field.type == expected_type

    def it_returns_empty_type_for_empty_instruction(self):
        from docx.oxml.parser import OxmlElement

        fldSimple = cast(CT_FldSimple, OxmlElement("w:fldSimple"))
        field = Field.for_simple(fldSimple)
        assert field.type == ""

    def it_exposes_the_result_text(self):
        fldSimple = cast(
            CT_FldSimple,
            element('w:fldSimple{w:instr=PAGE}/w:r/w:t"3"'),
        )
        field = Field.for_simple(fldSimple)
        assert field.result_text == "3"

    def it_returns_empty_result_text_when_no_run(self):
        fldSimple = cast(CT_FldSimple, element('w:fldSimple{w:instr=PAGE}'))
        field = Field.for_simple(fldSimple)
        assert field.result_text == ""


class DescribeField_Complex:
    """Unit-test suite for `Field` wrapping a complex field begin-run."""

    def _build_complex_paragraph(
        self, instr: str, result_text: str | None = None
    ) -> tuple[CT_P, CT_R]:
        p = cast(CT_P, element("w:p"))
        begin_run = p.add_complex_field(instr, result_text)
        return p, begin_run

    def it_is_complex(self):
        _, begin_run = self._build_complex_paragraph("PAGE", "3")
        field = Field.for_complex(begin_run)
        assert field.is_complex is True

    def it_reads_the_instruction(self):
        _, begin_run = self._build_complex_paragraph("REF bookmark1 \\h", "See here")
        field = Field.for_complex(begin_run)
        assert field.instruction == "REF bookmark1 \\h"

    def it_reads_the_type(self):
        _, begin_run = self._build_complex_paragraph("NUMPAGES", "10")
        field = Field.for_complex(begin_run)
        assert field.type == "NUMPAGES"

    def it_reads_the_result_text(self):
        _, begin_run = self._build_complex_paragraph("PAGE", "42")
        field = Field.for_complex(begin_run)
        assert field.result_text == "42"

    def it_returns_empty_result_text_when_no_separate_marker(self):
        # -- build a field with only begin/instrText/end; no separate marker --
        p = cast(CT_P, element("w:p"))
        p.add_complex_field("PAGE")  # no result_text => 4 runs
        # -- remove the separate marker to leave only begin/instrText/end --
        from docx.oxml.ns import qn

        seps = p.xpath('.//w:fldChar[@w:fldCharType="separate"]')
        sep_run = seps[0].getparent()
        sep_run.getparent().remove(sep_run)
        begin_run = p.r_lst[0]
        field = Field.for_complex(begin_run)
        assert field.result_text == ""

    def it_returns_empty_result_text_when_omitted(self):
        _, begin_run = self._build_complex_paragraph("PAGE")
        field = Field.for_complex(begin_run)
        assert field.result_text == ""

    def it_concatenates_instruction_split_across_runs(self):
        # -- some producers split the instruction across multiple instrText runs --
        from docx.oxml.ns import qn
        from docx.oxml.parser import OxmlElement

        p = cast(CT_P, element("w:p"))

        r_begin = p.add_r()
        fld_begin = OxmlElement("w:fldChar")
        fld_begin.set(qn("w:fldCharType"), "begin")
        r_begin.append(fld_begin)

        r_i1 = p.add_r()
        i1 = OxmlElement("w:instrText")
        i1.text = "REF "
        r_i1.append(i1)

        r_i2 = p.add_r()
        i2 = OxmlElement("w:instrText")
        i2.text = "bookmark1"
        r_i2.append(i2)

        r_sep = p.add_r()
        fld_sep = OxmlElement("w:fldChar")
        fld_sep.set(qn("w:fldCharType"), "separate")
        r_sep.append(fld_sep)

        r_end = p.add_r()
        fld_end = OxmlElement("w:fldChar")
        fld_end.set(qn("w:fldCharType"), "end"),
        r_end.append(fld_end)

        field = Field.for_complex(r_begin)
        assert field.instruction == "REF bookmark1"
        assert field.type == "REF"


class DescribeField_resolve:
    """Unit-test suite for `Field.resolve()` cross-reference resolution."""

    def _doc_with_bookmark(
        self, bookmark_name: str, bookmark_text: str
    ):
        """Return a `Document` whose body contains a bookmark wrapping `bookmark_text`.

        The bookmark uses id=0 and sits in a single paragraph alongside a
        ``w:r/w:t`` run containing `bookmark_text`.
        """
        from docx.document import Document
        from docx.oxml.document import CT_Document

        # -- build: <w:body><w:p><w:bookmarkStart .../><w:r><w:t>...</w:t></w:r>
        #    <w:bookmarkEnd .../></w:p></w:body> --
        doc_elm = cast(
            CT_Document,
            element(
                f"w:document/w:body/w:p/("
                f"w:bookmarkStart{{w:id=0,w:name={bookmark_name}}}"
                f",w:r/w:t\"{bookmark_text}\""
                f",w:bookmarkEnd{{w:id=0}}"
                f")"
            ),
        )
        return Document(doc_elm, None)  # type: ignore[arg-type]

    def it_resolves_REF_to_bookmark_text_for_simple_field(self):
        from docx.oxml.ns import qn
        from docx.oxml.parser import OxmlElement

        doc = self._doc_with_bookmark("Ref1", "Hello")
        fldSimple = cast(CT_FldSimple, OxmlElement("w:fldSimple"))
        fldSimple.set(qn("w:instr"), "REF Ref1 \\h")
        field = Field.for_simple(fldSimple)

        assert field.resolve(doc) == "Hello"

    def it_resolves_REF_to_bookmark_text_for_complex_field(self):
        doc = self._doc_with_bookmark("Ref1", "Hello")
        p = cast(CT_P, element("w:p"))
        begin_run = p.add_complex_field("REF Ref1 \\h", "stale")
        field = Field.for_complex(begin_run)

        assert field.resolve(doc) == "Hello"

    def it_resolves_heading_style_REF_referencing_a_heading_bookmark(self):
        # -- Word uses names like "_Ref12345" or "_Toc12345" for auto-generated
        #    cross-references; treat them like any other bookmark. --
        doc = self._doc_with_bookmark("_Ref12345", "Chapter 2")
        p = cast(CT_P, element("w:p"))
        begin_run = p.add_complex_field("REF _Ref12345 \\h", "")
        field = Field.for_complex(begin_run)

        assert field.resolve(doc) == "Chapter 2"

    def it_returns_cached_result_for_PAGEREF_when_present(self):
        doc = self._doc_with_bookmark("Ref1", "ignored")
        p = cast(CT_P, element("w:p"))
        begin_run = p.add_complex_field("PAGEREF Ref1 \\h", "42")
        field = Field.for_complex(begin_run)

        assert field.resolve(doc) == "42"

    def it_returns_question_mark_for_PAGEREF_when_no_cached_result(self):
        doc = self._doc_with_bookmark("Ref1", "ignored")
        p = cast(CT_P, element("w:p"))
        begin_run = p.add_complex_field("PAGEREF Ref1 \\h")
        field = Field.for_complex(begin_run)

        assert field.resolve(doc) == "?"

    def it_returns_cached_result_for_unrelated_field_types(self):
        doc = self._doc_with_bookmark("Ref1", "ignored")
        p = cast(CT_P, element("w:p"))
        begin_run = p.add_complex_field("PAGE", "7")
        field = Field.for_complex(begin_run)

        assert field.resolve(doc) == "7"

    def it_returns_cached_result_when_REF_bookmark_not_found(self):
        doc = self._doc_with_bookmark("Ref1", "Hello")
        p = cast(CT_P, element("w:p"))
        begin_run = p.add_complex_field("REF Missing \\h", "stale")
        field = Field.for_complex(begin_run)

        # -- unresolvable references leave the cached result untouched --
        assert field.resolve(doc) == "stale"

    def it_returns_cached_result_when_REF_has_no_bookmark_name(self):
        doc = self._doc_with_bookmark("Ref1", "Hello")
        p = cast(CT_P, element("w:p"))
        begin_run = p.add_complex_field("REF", "stale")
        field = Field.for_complex(begin_run)

        assert field.resolve(doc) == "stale"

    def it_strips_switches_before_looking_up_the_bookmark_name(self):
        # -- verify backslash switches are skipped --
        doc = self._doc_with_bookmark("Ref1", "Hello")
        p = cast(CT_P, element("w:p"))
        begin_run = p.add_complex_field('REF \\h Ref1 \\* MERGEFORMAT', "")
        field = Field.for_complex(begin_run)

        assert field.resolve(doc) == "Hello"

    def it_updates_result_text_in_place_for_simple_field(self):
        from docx.oxml.ns import qn
        from docx.oxml.parser import OxmlElement

        fldSimple = cast(CT_FldSimple, OxmlElement("w:fldSimple"))
        fldSimple.set(qn("w:instr"), "REF Ref1")
        r = fldSimple.add_r()
        r.add_t("old")
        field = Field.for_simple(fldSimple)

        field.update_result_text("new value")

        assert field.result_text == "new value"

    def it_updates_result_text_in_place_for_complex_field(self):
        p = cast(CT_P, element("w:p"))
        begin_run = p.add_complex_field("REF Ref1", "old")
        field = Field.for_complex(begin_run)

        field.update_result_text("new value")

        assert field.result_text == "new value"

    def it_preserves_whitespace_when_updated_text_has_leading_or_trailing_spaces(self):
        p = cast(CT_P, element("w:p"))
        begin_run = p.add_complex_field("REF Ref1", "old")
        field = Field.for_complex(begin_run)

        field.update_result_text(" padded ")

        assert field.result_text == " padded "

    def it_walks_bookmark_text_across_multiple_runs(self):
        # -- bookmark range spans two runs; concatenate their text --
        from docx.document import Document
        from docx.oxml.document import CT_Document

        doc_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/w:p/("
                "w:bookmarkStart{w:id=0,w:name=Ref1}"
                ',w:r/w:t"Hello "'
                ',w:r/w:t"World"'
                ",w:bookmarkEnd{w:id=0}"
                ")"
            ),
        )
        doc = Document(doc_elm, None)  # type: ignore[arg-type]
        p = cast(CT_P, element("w:p"))
        begin_run = p.add_complex_field("REF Ref1", "")
        field = Field.for_complex(begin_run)

        assert field.resolve(doc) == "Hello World"


class DescribeDocument_resolve_cross_references:
    """Unit-test suite for `Document.resolve_cross_references()`."""

    def it_resolves_REF_and_updates_result_text_in_place(self):
        from docx.document import Document
        from docx.oxml.document import CT_Document

        # -- body contains: a paragraph with bookmark "Ref1" wrapping "Hello",
        #    then a second paragraph with a REF simple field. --
        doc_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/("
                "w:p/("
                "w:bookmarkStart{w:id=0,w:name=Ref1}"
                ',w:r/w:t"Hello"'
                ",w:bookmarkEnd{w:id=0}"
                "),"
                "w:p/w:fldSimple{w:instr=REF Ref1}/w:r/w:t\"stale\""
                ")"
            ),
        )
        doc = Document(doc_elm, None)  # type: ignore[arg-type]

        count = doc.resolve_cross_references()

        assert count == 1
        # -- find the fldSimple and confirm its text is now "Hello" --
        fldSimples = doc._element.body.xpath(".//w:fldSimple")  # pyright: ignore[reportPrivateUsage]
        assert len(fldSimples) == 1
        assert Field.for_simple(fldSimples[0]).result_text == "Hello"

    def it_updates_multiple_fields_and_returns_the_count(self):
        from docx.document import Document
        from docx.oxml.document import CT_Document

        doc_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/("
                "w:p/("
                "w:bookmarkStart{w:id=0,w:name=A}"
                ',w:r/w:t"AAA"'
                ",w:bookmarkEnd{w:id=0}"
                "),"
                "w:p/("
                "w:bookmarkStart{w:id=1,w:name=B}"
                ',w:r/w:t"BBB"'
                ",w:bookmarkEnd{w:id=1}"
                "),"
                'w:p/w:fldSimple{w:instr=REF A}/w:r/w:t"stale"'
                ','
                'w:p/w:fldSimple{w:instr=REF B}/w:r/w:t"stale"'
                ")"
            ),
        )
        doc = Document(doc_elm, None)  # type: ignore[arg-type]

        count = doc.resolve_cross_references()

        assert count == 2
        results = [
            Field.for_simple(fs).result_text
            for fs in doc._element.body.xpath(".//w:fldSimple")  # pyright: ignore[reportPrivateUsage]
        ]
        assert results == ["AAA", "BBB"]

    def it_leaves_unresolvable_fields_untouched(self):
        from docx.document import Document
        from docx.oxml.document import CT_Document

        doc_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/"
                'w:p/w:fldSimple{w:instr=REF NoSuchBookmark}/w:r/w:t"stale"'
            ),
        )
        doc = Document(doc_elm, None)  # type: ignore[arg-type]

        count = doc.resolve_cross_references()

        assert count == 0
        fs = doc._element.body.xpath(".//w:fldSimple")[0]  # pyright: ignore[reportPrivateUsage]
        assert Field.for_simple(fs).result_text == "stale"

    def it_ignores_non_crossreference_fields(self):
        from docx.document import Document
        from docx.oxml.document import CT_Document

        doc_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/"
                'w:p/w:fldSimple{w:instr=PAGE}/w:r/w:t"3"'
            ),
        )
        doc = Document(doc_elm, None)  # type: ignore[arg-type]

        count = doc.resolve_cross_references()

        assert count == 0
        fs = doc._element.body.xpath(".//w:fldSimple")[0]  # pyright: ignore[reportPrivateUsage]
        assert Field.for_simple(fs).result_text == "3"

    def it_resolves_PAGEREF_with_empty_cached_result_to_question_mark(self):
        from docx.document import Document
        from docx.oxml.document import CT_Document

        # -- PAGEREF with no cached result_text; expect "?" after resolve --
        doc_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/("
                "w:p/("
                "w:bookmarkStart{w:id=0,w:name=Ref1}"
                ',w:r/w:t"X"'
                ",w:bookmarkEnd{w:id=0}"
                "),"
                "w:p/w:fldSimple{w:instr=PAGEREF Ref1}"
                ")"
            ),
        )
        doc = Document(doc_elm, None)  # type: ignore[arg-type]

        count = doc.resolve_cross_references()

        assert count == 1
        fs = [
            fs
            for fs in doc._element.body.xpath(".//w:fldSimple")  # pyright: ignore[reportPrivateUsage]
            if fs.get(qn("w:instr")) == "PAGEREF Ref1"
        ][0]
        assert Field.for_simple(fs).result_text == "?"


class DescribeField_DocPropertyResolution:
    """Closes upstream#1482 — resolve DOCPROPERTY / AUTHOR / TITLE / SUBJECT."""

    def _make_doc(self, body_cxml: str):
        from docx.document import Document
        from docx.oxml.document import CT_Document

        doc_elm = cast(CT_Document, element(body_cxml))
        return Document(doc_elm, None)  # type: ignore[arg-type]

    def it_resolves_AUTHOR_from_core_properties(self, request: pytest.FixtureRequest):
        from unittest.mock import PropertyMock

        from docx.opc.coreprops import CoreProperties

        doc = self._make_doc(
            'w:document/w:body/w:p/w:fldSimple{w:instr=AUTHOR}/w:r/w:t"OLD"'
        )
        core = CoreProperties(None)  # type: ignore[arg-type]
        pm = PropertyMock(return_value="Ada Lovelace")
        type(core).author = pm  # pyright: ignore[reportAttributeAccessIssue]
        request.addfinalizer(
            lambda: delattr(type(core), "author")
            if "author" in type(core).__dict__
            else None
        )

        fs = doc._element.body.xpath(".//w:fldSimple")[0]  # pyright: ignore[reportPrivateUsage]

        class _FakePart:
            core_properties = core

        doc._part = _FakePart()  # type: ignore[assignment]
        assert Field.for_simple(fs).resolve(doc) == "Ada Lovelace"

    def it_resolves_DOCPROPERTY_Author(self):
        from unittest.mock import PropertyMock
        from docx.opc.coreprops import CoreProperties

        doc = self._make_doc(
            "w:document/w:body/w:p/w:fldSimple"
            '{w:instr=DOCPROPERTY Author}/w:r/w:t"OLD"'
        )
        core = CoreProperties(None)  # type: ignore[arg-type]
        type(core).author = PropertyMock(return_value="Grace Hopper")

        class _FakePart:
            core_properties = core

        doc._part = _FakePart()  # type: ignore[assignment]
        fs = doc._element.body.xpath(".//w:fldSimple")[0]  # pyright: ignore[reportPrivateUsage]
        assert Field.for_simple(fs).resolve(doc) == "Grace Hopper"

    def it_resolves_DOCPROPERTY_custom_name(self):
        doc = self._make_doc(
            "w:document/w:body/w:p/w:fldSimple"
            '{w:instr=DOCPROPERTY Project}/w:r/w:t"OLD"'
        )

        class _FakeCustom:
            def get(self, name, default=None):
                return {"Project": "Apollo 11"}.get(name, default)

        class _FakeCore:
            pass

        class _FakePart:
            core_properties = _FakeCore()
            custom_properties = _FakeCustom()

        doc._part = _FakePart()  # type: ignore[assignment]
        fs = doc._element.body.xpath(".//w:fldSimple")[0]  # pyright: ignore[reportPrivateUsage]
        assert Field.for_simple(fs).resolve(doc) == "Apollo 11"

    def it_parses_quoted_docproperty_name(self):
        from docx.fields import _parse_docproperty_name

        assert _parse_docproperty_name('DOCPROPERTY "Some Multi Word"') == (
            "Some Multi Word"
        )
        assert _parse_docproperty_name("DOCPROPERTY Title") == "Title"
        assert _parse_docproperty_name("DOCPROPERTY") is None
        assert _parse_docproperty_name(
            'DOCPROPERTY "Simple" \\* MERGEFORMAT'
        ) == "Simple"


class DescribeField_MarkDirty:
    """Unit-test suite for `Field.mark_dirty` / `Field.is_dirty`."""

    def it_marks_a_simple_field_dirty(self):
        from docx.oxml.parser import OxmlElement

        fldSimple = cast(CT_FldSimple, OxmlElement("w:fldSimple"))
        fldSimple.set(qn("w:instr"), "PAGE")
        field = Field.for_simple(fldSimple)

        field.mark_dirty()

        assert fldSimple.get(qn("w:dirty")) == "true"
        assert field.is_dirty is True

    def it_reads_an_unset_simple_dirty_flag_as_false(self):
        from docx.oxml.parser import OxmlElement

        fldSimple = cast(CT_FldSimple, OxmlElement("w:fldSimple"))
        fldSimple.set(qn("w:instr"), "PAGE")
        field = Field.for_simple(fldSimple)
        assert field.is_dirty is False

    def it_marks_a_complex_field_dirty_on_the_begin_marker(self):
        p = cast(CT_P, element("w:p"))
        begin_run = p.add_complex_field(" TOC ", "cached")
        field = Field.for_complex(begin_run)

        field.mark_dirty()

        begin_fldChar = begin_run.find(qn("w:fldChar"))
        assert begin_fldChar is not None
        assert begin_fldChar.get(qn("w:dirty")) == "true"
        assert field.is_dirty is True

    def it_reads_an_unset_complex_dirty_flag_as_false(self):
        p = cast(CT_P, element("w:p"))
        begin_run = p.add_complex_field(" TOC ", "cached")
        assert Field.for_complex(begin_run).is_dirty is False

    @pytest.mark.parametrize(
        ("value", "expected"),
        [("true", True), ("1", True), ("on", True), ("false", False), ("0", False)],
    )
    def it_parses_known_dirty_values_consistently(
        self, value: str, expected: bool
    ):
        from docx.oxml.parser import OxmlElement

        fldSimple = cast(CT_FldSimple, OxmlElement("w:fldSimple"))
        fldSimple.set(qn("w:instr"), "PAGE")
        fldSimple.set(qn("w:dirty"), value)
        assert Field.for_simple(fldSimple).is_dirty is expected


class DescribeField_evaluate:
    """Unit-test suite for `Field.evaluate()` complex-field evaluation."""

    def _simple(self, instr: str, cached: str = "") -> Field:
        from docx.oxml.parser import OxmlElement

        fldSimple = cast(CT_FldSimple, OxmlElement("w:fldSimple"))
        fldSimple.set(qn("w:instr"), instr)
        if cached:
            r = OxmlElement("w:r")
            t = OxmlElement("w:t")
            t.text = cached
            r.append(t)
            fldSimple.append(r)
        return Field.for_simple(fldSimple)

    # -- MERGEFIELD ----------------------------------------------------------

    def it_evaluates_MERGEFIELD_from_context(self):
        field = self._simple("MERGEFIELD firstname", cached="<<firstname>>")
        assert field.evaluate({"firstname": "Ada"}) == "Ada"

    def it_falls_back_to_cached_result_when_MERGEFIELD_key_missing(self):
        field = self._simple("MERGEFIELD firstname", cached="<<firstname>>")
        assert field.evaluate({}) == "<<firstname>>"

    def it_evaluates_MERGEFIELD_with_quoted_multiword_name(self):
        field = self._simple('MERGEFIELD "Full Name"')
        assert field.evaluate({"Full Name": "Ada Lovelace"}) == "Ada Lovelace"

    # -- IF ------------------------------------------------------------------

    def it_evaluates_IF_with_equal_match(self):
        field = self._simple('IF "yes" = "yes" "match" "nope"')
        assert field.evaluate({}) == "match"

    def it_evaluates_IF_with_equal_mismatch(self):
        field = self._simple('IF "yes" = "no" "match" "nope"')
        assert field.evaluate({}) == "nope"

    def it_evaluates_IF_with_nested_MERGEFIELD(self):
        field = self._simple('IF {MERGEFIELD status} = "active" "OK" "FAIL"')
        assert field.evaluate({"status": "active"}) == "OK"
        assert field.evaluate({"status": "cancelled"}) == "FAIL"

    @pytest.mark.parametrize(
        ("op", "lhs", "rhs", "expected"),
        [
            ("<>", "a", "b", "yes"),
            ("<>", "a", "a", "no"),
            ("!=", "a", "a", "no"),
            ("<", "1", "2", "yes"),
            (">", "2", "1", "yes"),
            ("<=", "2", "2", "yes"),
            (">=", "2", "3", "no"),
        ],
    )
    def it_supports_the_common_comparison_operators(
        self, op: str, lhs: str, rhs: str, expected: str
    ):
        field = self._simple(f'IF "{lhs}" {op} "{rhs}" "yes" "no"')
        assert field.evaluate({}) == expected

    def it_uses_numeric_comparison_when_both_sides_parse_as_numbers(self):
        field = self._simple('IF "10" > "9" "big" "small"')
        assert field.evaluate({}) == "big"

    def it_returns_empty_false_branch_when_only_true_text_given(self):
        field = self._simple('IF "a" = "b" "match"')
        assert field.evaluate({}) == ""

    def it_returns_cached_result_when_IF_is_malformed(self):
        field = self._simple("IF", cached="x")
        assert field.evaluate({}) == "x"

    # -- HYPERLINK -----------------------------------------------------------

    def it_returns_cached_display_text_for_HYPERLINK_when_present(self):
        field = self._simple('HYPERLINK "https://example.com"', cached="click me")
        assert field.evaluate({}) == "click me"

    def it_returns_the_url_for_HYPERLINK_when_no_cached_text(self):
        field = self._simple('HYPERLINK "https://example.com"')
        assert field.evaluate({}) == "https://example.com"

    # -- runtime-dynamic -----------------------------------------------------

    @pytest.mark.parametrize(
        "instr", ["PAGE", "NUMPAGES", "DATE", "TIME"]
    )
    def it_returns_cached_result_for_runtime_dynamic_fields(self, instr: str):
        field = self._simple(instr, cached="7")
        assert field.evaluate({}) == "7"

    @pytest.mark.parametrize(
        "instr", ["PAGE", "NUMPAGES", "DATE", "TIME"]
    )
    def it_returns_question_mark_for_runtime_dynamic_fields_with_no_cache(
        self, instr: str
    ):
        field = self._simple(instr)
        assert field.evaluate({}) == "?"

    # -- formula (=) ---------------------------------------------------------

    @pytest.mark.parametrize(
        ("expr", "expected"),
        [
            ("= 1+2", "3"),
            ("= 2*3", "6"),
            ("= (2+3)*4", "20"),
            ("= 7/2", "3.5"),
            ("= 10 % 3", "1"),
            ("= 2**8", "256"),
        ],
    )
    def it_evaluates_arithmetic_formula_fields(self, expr: str, expected: str):
        field = self._simple(expr)
        assert field.evaluate({}) == expected

    def it_substitutes_MERGEFIELD_in_formula(self):
        field = self._simple("= {MERGEFIELD qty} * 10")
        assert field.evaluate({"qty": 4}) == "40"

    def it_returns_cached_for_formula_with_disallowed_chars(self):
        field = self._simple('= __import__("os").system("ls")', cached="stale")
        assert field.evaluate({}) == "stale"

    # -- pass-through --------------------------------------------------------

    def it_returns_cached_result_for_unknown_field_types(self):
        field = self._simple("TOC \\o \"1-3\"", cached="toc-preview")
        assert field.evaluate({}) == "toc-preview"


class DescribeDocument_evaluate_fields:
    """Unit-test suite for `Document.evaluate_fields()`."""

    def it_updates_MERGEFIELD_cached_text_in_place(self):
        from docx.document import Document
        from docx.oxml.document import CT_Document

        doc_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/"
                'w:p/w:fldSimple{w:instr=MERGEFIELD name}/w:r/w:t"stale"'
            ),
        )
        doc = Document(doc_elm, None)  # type: ignore[arg-type]
        count = doc.evaluate_fields({"name": "Ada"})
        assert count == 1
        fs = doc._element.body.xpath(".//w:fldSimple")[0]  # pyright: ignore[reportPrivateUsage]
        assert Field.for_simple(fs).result_text == "Ada"

    def it_returns_zero_when_nothing_changed(self):
        from docx.document import Document
        from docx.oxml.document import CT_Document

        doc_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/"
                'w:p/w:fldSimple{w:instr=MERGEFIELD name}/w:r/w:t"Ada"'
            ),
        )
        doc = Document(doc_elm, None)  # type: ignore[arg-type]
        count = doc.evaluate_fields({"name": "Ada"})
        assert count == 0

    def it_evaluates_IF_formula_and_MERGEFIELD_together(self):
        from docx.document import Document
        from docx.oxml.document import CT_Document
        from docx.oxml.parser import OxmlElement

        doc_elm = cast(
            CT_Document,
            element("w:document/w:body/(w:p,w:p,w:p)"),
        )
        ps = doc_elm.body.xpath(".//w:p")
        # -- merge field --
        fs1 = OxmlElement("w:fldSimple")
        fs1.set(qn("w:instr"), "MERGEFIELD name")
        ps[0].append(fs1)
        # -- IF --
        fs2 = OxmlElement("w:fldSimple")
        fs2.set(qn("w:instr"), 'IF {MERGEFIELD status} = "active" "yes" "no"')
        ps[1].append(fs2)
        # -- formula --
        fs3 = OxmlElement("w:fldSimple")
        fs3.set(qn("w:instr"), "= 2+3")
        ps[2].append(fs3)

        doc = Document(doc_elm, None)  # type: ignore[arg-type]
        count = doc.evaluate_fields({"name": "Ada", "status": "active"})
        assert count == 3
        fs_list = doc._element.body.xpath(".//w:fldSimple")  # pyright: ignore[reportPrivateUsage]
        assert [Field.for_simple(f).result_text for f in fs_list] == [
            "Ada", "yes", "5"
        ]

    def it_passes_document_context_for_property_fields(
        self, request: pytest.FixtureRequest
    ):
        from docx.document import Document
        from docx.oxml.document import CT_Document

        doc_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/"
                'w:p/w:fldSimple{w:instr=AUTHOR}/w:r/w:t"stale"'
            ),
        )
        doc = Document(doc_elm, None)  # type: ignore[arg-type]

        # -- stub core_properties.author --
        class _Props:
            author = "Jane Doe"
            title = None
            subject = None
            keywords = None
            comments = None
            last_modified_by = None

        doc.__class__.core_properties = property(  # type: ignore[assignment]
            lambda self: _Props()
        )
        try:
            count = doc.evaluate_fields({})
            assert count == 1
            fs = doc._element.body.xpath(".//w:fldSimple")[0]  # pyright: ignore[reportPrivateUsage]
            assert Field.for_simple(fs).result_text == "Jane Doe"
        finally:
            del doc.__class__.core_properties  # type: ignore[attr-defined]

    def it_accepts_a_missing_context_argument(self):
        from docx.document import Document
        from docx.oxml.document import CT_Document

        doc_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/"
                'w:p/w:fldSimple{w:instr=PAGE}/w:r/w:t"7"'
            ),
        )
        doc = Document(doc_elm, None)  # type: ignore[arg-type]
        count = doc.evaluate_fields()
        # -- PAGE with cached "7" stays "7" --
        assert count == 0
