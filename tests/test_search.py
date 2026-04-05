# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.search` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.document import Document
from docx.oxml.document import CT_Document
from docx.parts.document import DocumentPart
from docx.search import (
    SearchMatch,
    _build_char_map,
    _find_matches_in_text,
    _replace_in_paragraph,
    _run_indices_for_match,
    replace_in_paragraphs,
    search_paragraphs,
)

from .unitutil.cxml import element
from .unitutil.mock import FixtureRequest, Mock, instance_mock


class DescribeSearchMatch:
    """Unit-test suite for `docx.search.SearchMatch`."""

    def it_stores_match_attributes(self):
        paragraph_ = Mock()
        match = SearchMatch(
            paragraph=paragraph_,
            paragraph_index=2,
            run_indices=[0, 1],
            start=5,
            end=10,
        )

        assert match.paragraph is paragraph_
        assert match.paragraph_index == 2
        assert match.run_indices == [0, 1]
        assert match.start == 5
        assert match.end == 10


class Describe_build_char_map:
    """Unit-test suite for `docx.search._build_char_map`."""

    def it_builds_a_char_map_for_a_single_run(self, fake_parent: Mock):
        p = element("w:p/w:r/w:t\"hello\"")
        from docx.text.paragraph import Paragraph

        para = Paragraph(p, fake_parent)
        runs = para.runs
        full_text, char_map = _build_char_map(runs)

        assert full_text == "hello"
        assert len(char_map) == 5
        assert all(cm.run_index == 0 for cm in char_map)

    def it_builds_a_char_map_for_multiple_runs(self, fake_parent: Mock):
        p = element('w:p/(w:r/w:t"hel",w:r/w:t"lo")')
        from docx.text.paragraph import Paragraph

        para = Paragraph(p, fake_parent)
        runs = para.runs
        full_text, char_map = _build_char_map(runs)

        assert full_text == "hello"
        assert len(char_map) == 5
        # first 3 chars in run 0, last 2 in run 1
        assert [cm.run_index for cm in char_map] == [0, 0, 0, 1, 1]

    def it_returns_empty_for_no_runs(self):
        full_text, char_map = _build_char_map([])

        assert full_text == ""
        assert char_map == []


class Describe_find_matches_in_text:
    """Unit-test suite for `docx.search._find_matches_in_text`."""

    def it_finds_simple_matches(self):
        matches = _find_matches_in_text("hello world hello", "hello", True, False)
        assert matches == [(0, 5), (12, 17)]

    def it_finds_case_insensitive_matches(self):
        matches = _find_matches_in_text("Hello HELLO hello", "hello", False, False)
        assert matches == [(0, 5), (6, 11), (12, 17)]

    def it_finds_whole_word_matches(self):
        matches = _find_matches_in_text("hello helloworld hello", "hello", True, True)
        assert matches == [(0, 5), (17, 22)]

    def it_returns_empty_for_no_matches(self):
        matches = _find_matches_in_text("hello world", "xyz", True, False)
        assert matches == []

    def it_escapes_regex_special_characters(self):
        matches = _find_matches_in_text("price is $10.00", "$10.00", True, False)
        assert matches == [(9, 15)]


class Describe_run_indices_for_match:
    """Unit-test suite for `docx.search._run_indices_for_match`."""

    def it_returns_indices_for_single_run_match(self, fake_parent: Mock):
        p = element('w:p/(w:r/w:t"hello",w:r/w:t" world")')
        from docx.text.paragraph import Paragraph

        para = Paragraph(p, fake_parent)
        _, char_map = _build_char_map(para.runs)

        indices = _run_indices_for_match(char_map, 0, 5)
        assert indices == [0]

    def it_returns_indices_for_multi_run_match(self, fake_parent: Mock):
        p = element('w:p/(w:r/w:t"hel",w:r/w:t"lo w",w:r/w:t"orld")')
        from docx.text.paragraph import Paragraph

        para = Paragraph(p, fake_parent)
        _, char_map = _build_char_map(para.runs)

        # "hello" spans runs 0 and 1
        indices = _run_indices_for_match(char_map, 0, 5)
        assert indices == [0, 1]

    def it_returns_empty_for_empty_char_map(self):
        assert _run_indices_for_match([], 0, 5) == []


class DescribeSearchParagraphs:
    """Unit-test suite for `docx.search.search_paragraphs`."""

    def it_finds_matches_across_paragraphs(self, fake_parent: Mock):
        from docx.text.paragraph import Paragraph

        p1 = element('w:p/w:r/w:t"hello world"')
        p2 = element('w:p/w:r/w:t"say hello"')
        paragraphs = [Paragraph(p1, fake_parent), Paragraph(p2, fake_parent)]

        matches = search_paragraphs(paragraphs, "hello")

        assert len(matches) == 2
        assert matches[0].paragraph_index == 0
        assert matches[0].start == 0
        assert matches[0].end == 5
        assert matches[1].paragraph_index == 1
        assert matches[1].start == 4
        assert matches[1].end == 9

    def it_finds_matches_spanning_runs(self, fake_parent: Mock):
        from docx.text.paragraph import Paragraph

        p = element('w:p/(w:r/w:t"hel",w:r/w:t"lo")')
        paragraphs = [Paragraph(p, fake_parent)]

        matches = search_paragraphs(paragraphs, "hello")

        assert len(matches) == 1
        assert matches[0].run_indices == [0, 1]

    def it_returns_empty_for_no_matches(self, fake_parent: Mock):
        from docx.text.paragraph import Paragraph

        p = element('w:p/w:r/w:t"hello"')
        paragraphs = [Paragraph(p, fake_parent)]

        matches = search_paragraphs(paragraphs, "xyz")
        assert matches == []

    def it_skips_paragraphs_with_no_runs(self, fake_parent: Mock):
        from docx.text.paragraph import Paragraph

        p = element("w:p")
        paragraphs = [Paragraph(p, fake_parent)]

        matches = search_paragraphs(paragraphs, "hello")
        assert matches == []


class DescribeReplaceSingleRun:
    """Unit-test suite for replacement within a single run."""

    def it_replaces_text_in_a_single_run(self, fake_parent: Mock):
        from docx.text.paragraph import Paragraph

        p = element('w:p/w:r/w:t"hello world"')
        para = Paragraph(p, fake_parent)

        count = _replace_in_paragraph(para, "hello", "goodbye")

        assert count == 1
        assert para.runs[0].text == "goodbye world"

    def it_replaces_multiple_occurrences(self, fake_parent: Mock):
        from docx.text.paragraph import Paragraph

        p = element('w:p/w:r/w:t"ab ab ab"')
        para = Paragraph(p, fake_parent)

        count = _replace_in_paragraph(para, "ab", "cd")

        assert count == 3
        assert para.runs[0].text == "cd cd cd"

    def it_returns_zero_for_no_matches(self, fake_parent: Mock):
        from docx.text.paragraph import Paragraph

        p = element('w:p/w:r/w:t"hello"')
        para = Paragraph(p, fake_parent)

        count = _replace_in_paragraph(para, "xyz", "abc")

        assert count == 0
        assert para.runs[0].text == "hello"


class DescribeReplaceMultiRun:
    """Unit-test suite for replacement spanning multiple runs."""

    def it_replaces_text_spanning_two_runs(self, fake_parent: Mock):
        from docx.text.paragraph import Paragraph

        p = element('w:p/(w:r/w:t"hel",w:r/w:t"lo world")')
        para = Paragraph(p, fake_parent)

        count = _replace_in_paragraph(para, "hello", "hi")

        assert count == 1
        # replacement goes in first run, remainder of second run preserved
        assert para.runs[0].text == "hi"
        assert para.runs[1].text == " world"

    def it_replaces_text_spanning_three_runs(self, fake_parent: Mock):
        from docx.text.paragraph import Paragraph

        p = element('w:p/(w:r/w:t"he",w:r/w:t"ll",w:r/w:t"o world")')
        para = Paragraph(p, fake_parent)

        count = _replace_in_paragraph(para, "hello", "hi")

        assert count == 1
        assert para.runs[0].text == "hi"
        assert para.runs[1].text == ""
        assert para.runs[2].text == " world"

    def it_preserves_formatting_of_first_run(self, fake_parent: Mock):
        from docx.text.paragraph import Paragraph

        # first run is bold, second is not
        p = element('w:p/(w:r/(w:rPr/w:b,w:t"hel"),w:r/w:t"lo")')
        para = Paragraph(p, fake_parent)

        _replace_in_paragraph(para, "hello", "goodbye")

        # the first run should still be bold (rPr preserved by run.text setter)
        assert para.runs[0].bold is True
        assert para.runs[0].text == "goodbye"
        assert para.runs[1].text == ""


class DescribeReplaceInParagraphs:
    """Unit-test suite for `docx.search.replace_in_paragraphs`."""

    def it_replaces_across_multiple_paragraphs(self, fake_parent: Mock):
        from docx.text.paragraph import Paragraph

        p1 = element('w:p/w:r/w:t"hello world"')
        p2 = element('w:p/w:r/w:t"say hello"')
        paragraphs = [Paragraph(p1, fake_parent), Paragraph(p2, fake_parent)]

        count = replace_in_paragraphs(paragraphs, "hello", "hi")

        assert count == 2
        assert paragraphs[0].runs[0].text == "hi world"
        assert paragraphs[1].runs[0].text == "say hi"

    def it_supports_case_insensitive_replace(self, fake_parent: Mock):
        from docx.text.paragraph import Paragraph

        p = element('w:p/w:r/w:t"Hello HELLO hello"')
        paragraphs = [Paragraph(p, fake_parent)]

        count = replace_in_paragraphs(paragraphs, "hello", "hi", case_sensitive=False)

        assert count == 3
        assert paragraphs[0].runs[0].text == "hi hi hi"

    def it_supports_whole_word_replace(self, fake_parent: Mock):
        from docx.text.paragraph import Paragraph

        p = element('w:p/w:r/w:t"hello helloworld hello"')
        paragraphs = [Paragraph(p, fake_parent)]

        count = replace_in_paragraphs(paragraphs, "hello", "hi", whole_word=True)

        assert count == 2
        assert paragraphs[0].runs[0].text == "hi helloworld hi"


class DescribeDocumentSearchAndReplace:
    """Unit-test suite for `Document.search()` and `Document.replace()`."""

    def it_can_search_the_document(self, document_part_: Mock):
        doc_elm = cast(
            CT_Document,
            element('w:document/w:body/(w:p/w:r/w:t"hello world",w:p/w:r/w:t"say hello")'),
        )
        document = Document(doc_elm, document_part_)

        matches = document.search("hello")

        assert len(matches) == 2
        assert isinstance(matches[0], SearchMatch)
        assert matches[0].paragraph_index == 0
        assert matches[1].paragraph_index == 1

    def it_can_replace_in_the_document(self, document_part_: Mock):
        doc_elm = cast(
            CT_Document,
            element('w:document/w:body/(w:p/w:r/w:t"hello world",w:p/w:r/w:t"say hello")'),
        )
        document = Document(doc_elm, document_part_)

        count = document.replace("hello", "hi")

        assert count == 2
        assert document.paragraphs[0].text == "hi world"
        assert document.paragraphs[1].text == "say hi"

    def it_can_search_case_insensitively(self, document_part_: Mock):
        doc_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"Hello HELLO"'),
        )
        document = Document(doc_elm, document_part_)

        matches = document.search("hello", case_sensitive=False)

        assert len(matches) == 2

    def it_can_search_whole_words(self, document_part_: Mock):
        doc_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"hello helloworld"'),
        )
        document = Document(doc_elm, document_part_)

        matches = document.search("hello", whole_word=True)

        assert len(matches) == 1
        assert matches[0].start == 0

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def document_part_(self, request: FixtureRequest):
        return instance_mock(request, DocumentPart)
