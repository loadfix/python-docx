# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.search` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.document import Document
from docx.oxml.document import CT_Document
from docx.search import (
    SearchMatch,
    _build_char_map,
    replace_in_paragraphs,
    replace_in_paragraphs_regex,
    search_paragraphs,
    search_paragraphs_regex,
)
from docx.text.paragraph import Paragraph

from .unitutil.cxml import element
from .unitutil.mock import Mock

import re


class DescribeSearchMatch:
    """Unit-test suite for `docx.search.SearchMatch` objects."""

    def it_provides_access_to_its_properties(self):
        paragraph_ = Mock(spec=Paragraph)
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


class DescribeSearch:
    """Unit-test suite for `docx.search.search_paragraphs`."""

    def it_finds_text_in_a_single_run(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"hello world"'),
        )
        doc = Document(document_elm, Mock())
        paragraphs = doc.paragraphs

        matches = search_paragraphs(paragraphs, "world")

        assert len(matches) == 1
        assert matches[0].paragraph_index == 0
        assert matches[0].start == 6
        assert matches[0].end == 11
        assert matches[0].run_indices == [0]

    def it_finds_text_spanning_multiple_runs(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/(w:r/w:t"hel",w:r/w:t"lo world")'),
        )
        doc = Document(document_elm, Mock())
        paragraphs = doc.paragraphs

        matches = search_paragraphs(paragraphs, "hello")

        assert len(matches) == 1
        assert matches[0].run_indices == [0, 1]
        assert matches[0].start == 0
        assert matches[0].end == 5

    def it_finds_multiple_matches_in_one_paragraph(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"foo bar foo"'),
        )
        doc = Document(document_elm, Mock())

        matches = search_paragraphs(doc.paragraphs, "foo")

        assert len(matches) == 2
        assert matches[0].start == 0
        assert matches[0].end == 3
        assert matches[1].start == 8
        assert matches[1].end == 11

    def it_finds_matches_across_multiple_paragraphs(self):
        document_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/"
                '(w:p/w:r/w:t"hello"'
                ',w:p/w:r/w:t"world"'
                ',w:p/w:r/w:t"hello again")'
            ),
        )
        doc = Document(document_elm, Mock())

        matches = search_paragraphs(doc.paragraphs, "hello")

        assert len(matches) == 2
        assert matches[0].paragraph_index == 0
        assert matches[1].paragraph_index == 2

    def it_returns_empty_list_when_no_match(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"hello"'),
        )
        doc = Document(document_elm, Mock())

        matches = search_paragraphs(doc.paragraphs, "xyz")

        assert matches == []

    def it_returns_empty_list_for_empty_search_text(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"hello"'),
        )
        doc = Document(document_elm, Mock())

        matches = search_paragraphs(doc.paragraphs, "")

        assert matches == []

    def it_supports_case_insensitive_search(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"Hello World"'),
        )
        doc = Document(document_elm, Mock())

        matches = search_paragraphs(doc.paragraphs, "hello", case_sensitive=False)

        assert len(matches) == 1
        assert matches[0].start == 0
        assert matches[0].end == 5

    def it_supports_case_sensitive_search_by_default(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"Hello World"'),
        )
        doc = Document(document_elm, Mock())

        matches = search_paragraphs(doc.paragraphs, "hello")

        assert matches == []

    def it_supports_whole_word_search(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"cat concatenate the cat"'),
        )
        doc = Document(document_elm, Mock())

        matches = search_paragraphs(doc.paragraphs, "cat", whole_word=True)

        assert len(matches) == 2
        assert matches[0].start == 0
        assert matches[0].end == 3
        assert matches[1].start == 20
        assert matches[1].end == 23

    def it_handles_paragraph_with_no_runs(self):
        document_elm = cast(
            CT_Document,
            element("w:document/w:body/w:p"),
        )
        doc = Document(document_elm, Mock())

        matches = search_paragraphs(doc.paragraphs, "text")

        assert matches == []


class DescribeReplace:
    """Unit-test suite for `docx.search.replace_in_paragraphs`."""

    def it_replaces_text_in_a_single_run(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"hello world"'),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs(doc.paragraphs, "world", "there")

        assert count == 1
        assert doc.paragraphs[0].text == "hello there"

    def it_replaces_text_spanning_multiple_runs(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/(w:r/w:t"hel",w:r/w:t"lo world")'),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs(doc.paragraphs, "hello", "hi")

        assert count == 1
        # First run gets the replacement text, second run loses the matched portion.
        assert doc.paragraphs[0].runs[0].text == "hi"
        assert doc.paragraphs[0].runs[1].text == " world"

    def it_replaces_multiple_occurrences(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"foo bar foo"'),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs(doc.paragraphs, "foo", "baz")

        assert count == 2
        assert doc.paragraphs[0].text == "baz bar baz"

    def it_replaces_across_multiple_paragraphs(self):
        document_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/"
                '(w:p/w:r/w:t"hello"'
                ',w:p/w:r/w:t"world"'
                ',w:p/w:r/w:t"hello")'
            ),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs(doc.paragraphs, "hello", "hi")

        assert count == 2
        assert doc.paragraphs[0].text == "hi"
        assert doc.paragraphs[1].text == "world"
        assert doc.paragraphs[2].text == "hi"

    def it_preserves_formatting_of_first_run(self):
        document_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/w:p/"
                "(w:r/(w:rPr/w:b,w:t\"hel\")"
                ",w:r/(w:rPr/w:i,w:t\"lo world\"))"
            ),
        )
        doc = Document(document_elm, Mock())

        replace_in_paragraphs(doc.paragraphs, "hello", "hi")

        # First run keeps its bold formatting.
        assert doc.paragraphs[0].runs[0].bold is True
        assert doc.paragraphs[0].runs[0].text == "hi"
        # Second run keeps its italic formatting.
        assert doc.paragraphs[0].runs[1].italic is True
        assert doc.paragraphs[0].runs[1].text == " world"

    def it_handles_replacement_with_longer_text(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"hi"'),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs(doc.paragraphs, "hi", "hello world")

        assert count == 1
        assert doc.paragraphs[0].text == "hello world"

    def it_handles_replacement_with_empty_text(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"hello world"'),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs(doc.paragraphs, "world", "")

        assert count == 1
        assert doc.paragraphs[0].text == "hello "

    def it_returns_zero_when_no_match(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"hello"'),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs(doc.paragraphs, "xyz", "abc")

        assert count == 0
        assert doc.paragraphs[0].text == "hello"

    def it_returns_zero_for_empty_old_text(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"hello"'),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs(doc.paragraphs, "", "abc")

        assert count == 0

    def it_supports_case_insensitive_replace(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"Hello HELLO hello"'),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs(
            doc.paragraphs, "hello", "hi", case_sensitive=False
        )

        assert count == 3
        assert doc.paragraphs[0].text == "hi hi hi"

    def it_supports_whole_word_replace(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"cat concatenate the cat"'),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs(
            doc.paragraphs, "cat", "dog", whole_word=True
        )

        assert count == 2
        assert doc.paragraphs[0].text == "dog concatenate the dog"

    def it_replaces_text_spanning_three_runs(self):
        document_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/w:p/"
                '(w:r/w:t"ab",w:r/w:t"cd",w:r/w:t"ef")'
            ),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs(doc.paragraphs, "bcde", "X")

        assert count == 1
        assert doc.paragraphs[0].runs[0].text == "aX"
        assert doc.paragraphs[0].runs[1].text == ""
        assert doc.paragraphs[0].runs[2].text == "f"


class DescribeDocumentSearchAndReplace:
    """Unit-test suite for Document.search() and Document.replace()."""

    def it_exposes_search_on_document(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"hello world"'),
        )
        doc = Document(document_elm, Mock())

        matches = doc.search("world")

        assert len(matches) == 1
        assert matches[0].start == 6

    def it_exposes_replace_on_document(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"hello world"'),
        )
        doc = Document(document_elm, Mock())

        count = doc.replace("world", "there")

        assert count == 1
        assert doc.paragraphs[0].text == "hello there"

    def it_passes_options_through_to_search(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"Hello HELLO"'),
        )
        doc = Document(document_elm, Mock())

        matches = doc.search("hello", case_sensitive=False)

        assert len(matches) == 2

    def it_passes_options_through_to_replace(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"cat concatenate"'),
        )
        doc = Document(document_elm, Mock())

        count = doc.replace("cat", "dog", whole_word=True)

        assert count == 1
        assert doc.paragraphs[0].text == "dog concatenate"


class DescribeSearchRegex:
    """Unit-test suite for `docx.search.search_paragraphs_regex`."""

    def it_finds_a_simple_regex_match(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"order 12345 shipped"'),
        )
        doc = Document(document_elm, Mock())

        matches = search_paragraphs_regex(doc.paragraphs, r"\d+")

        assert len(matches) == 1
        assert matches[0].start == 6
        assert matches[0].end == 11
        assert matches[0].run_indices == [0]

    def it_supports_case_insensitive_flag(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"Hello HELLO hello"'),
        )
        doc = Document(document_elm, Mock())

        matches = search_paragraphs_regex(doc.paragraphs, r"hello", re.IGNORECASE)

        assert len(matches) == 3

    def it_accepts_a_compiled_pattern(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"Hello HELLO"'),
        )
        doc = Document(document_elm, Mock())
        compiled = re.compile(r"hello", re.IGNORECASE)

        matches = search_paragraphs_regex(doc.paragraphs, compiled)

        assert len(matches) == 2

    def it_finds_multiple_matches_in_one_paragraph(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"a1 b2 c3 d4"'),
        )
        doc = Document(document_elm, Mock())

        matches = search_paragraphs_regex(doc.paragraphs, r"[a-z]\d")

        assert len(matches) == 4
        assert [m.start for m in matches] == [0, 3, 6, 9]

    def it_finds_matches_across_run_boundaries(self):
        document_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/w:p/"
                '(w:r/w:t"foo",w:r/w:t"BAR",w:r/w:t"baz")'
            ),
        )
        doc = Document(document_elm, Mock())

        matches = search_paragraphs_regex(doc.paragraphs, r"ooBARba")

        assert len(matches) == 1
        assert matches[0].start == 1
        assert matches[0].end == 8
        assert matches[0].run_indices == [0, 1, 2]

    def it_returns_correct_match_offsets_across_paragraphs(self):
        document_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/"
                '(w:p/w:r/w:t"hello 42"'
                ',w:p/w:r/w:t"world 99")'
            ),
        )
        doc = Document(document_elm, Mock())

        matches = search_paragraphs_regex(doc.paragraphs, r"\d+")

        assert len(matches) == 2
        assert matches[0].paragraph_index == 0
        assert matches[0].start == 6
        assert matches[0].end == 8
        assert matches[1].paragraph_index == 1
        assert matches[1].start == 6
        assert matches[1].end == 8

    def it_returns_empty_list_when_no_match(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"hello"'),
        )
        doc = Document(document_elm, Mock())

        matches = search_paragraphs_regex(doc.paragraphs, r"\d+")

        assert matches == []


class DescribeReplaceRegex:
    """Unit-test suite for `docx.search.replace_in_paragraphs_regex`."""

    def it_replaces_a_simple_regex_match(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"order 12345 shipped"'),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs_regex(doc.paragraphs, r"\d+", "N/A")

        assert count == 1
        assert doc.paragraphs[0].text == "order N/A shipped"

    def it_supports_case_insensitive_flag(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"Hello HELLO hello"'),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs_regex(
            doc.paragraphs, r"hello", "hi", re.IGNORECASE
        )

        assert count == 3
        assert doc.paragraphs[0].text == "hi hi hi"

    def it_expands_backreferences(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"foobar"'),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs_regex(doc.paragraphs, r"(foo)bar", r"\1baz")

        assert count == 1
        assert doc.paragraphs[0].text == "foobaz"

    def it_expands_named_backreferences(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"Mr. Smith"'),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs_regex(
            doc.paragraphs, r"Mr\. (?P<name>\w+)", r"Dr. \g<name>"
        )

        assert count == 1
        assert doc.paragraphs[0].text == "Dr. Smith"

    def it_accepts_a_compiled_pattern(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"Hello HELLO"'),
        )
        doc = Document(document_elm, Mock())
        compiled = re.compile(r"hello", re.IGNORECASE)

        count = replace_in_paragraphs_regex(doc.paragraphs, compiled, "hi")

        assert count == 2
        assert doc.paragraphs[0].text == "hi hi"

    def it_replaces_multiple_matches_in_one_paragraph(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"a1 b2 c3"'),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs_regex(doc.paragraphs, r"[a-z]\d", "X")

        assert count == 3
        assert doc.paragraphs[0].text == "X X X"

    def it_replaces_match_spanning_three_runs(self):
        document_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/w:p/"
                '(w:r/w:t"ab",w:r/w:t"cd",w:r/w:t"ef")'
            ),
        )
        doc = Document(document_elm, Mock())

        # Regex matches part of each of the three runs.
        count = replace_in_paragraphs_regex(doc.paragraphs, r"b.de", "X")

        assert count == 1
        assert doc.paragraphs[0].runs[0].text == "aX"
        assert doc.paragraphs[0].runs[1].text == ""
        assert doc.paragraphs[0].runs[2].text == "f"

    def it_preserves_first_run_formatting_on_cross_run_match(self):
        document_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/w:p/"
                "(w:r/(w:rPr/w:b,w:t\"hel\")"
                ",w:r/(w:rPr/w:i,w:t\"lo world\"))"
            ),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs_regex(doc.paragraphs, r"h\w+o", "hi")

        assert count == 1
        assert doc.paragraphs[0].runs[0].bold is True
        assert doc.paragraphs[0].runs[0].text == "hi"
        assert doc.paragraphs[0].runs[1].italic is True
        assert doc.paragraphs[0].runs[1].text == " world"

    def it_returns_zero_when_no_match(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"hello"'),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs_regex(doc.paragraphs, r"\d+", "N/A")

        assert count == 0
        assert doc.paragraphs[0].text == "hello"

    def it_handles_paragraph_with_no_runs(self):
        document_elm = cast(
            CT_Document,
            element("w:document/w:body/w:p"),
        )
        doc = Document(document_elm, Mock())

        count = replace_in_paragraphs_regex(doc.paragraphs, r"\d+", "N/A")

        assert count == 0

    def it_skips_zero_width_matches(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"abc"'),
        )
        doc = Document(document_elm, Mock())

        # ``\b`` is a zero-width assertion. Nothing should be replaced.
        count = replace_in_paragraphs_regex(doc.paragraphs, r"\b", "X")

        assert count == 0
        assert doc.paragraphs[0].text == "abc"


class DescribeDocumentRegex:
    """Unit-test suite for Document.search_regex() and Document.replace_regex()."""

    def it_exposes_search_regex_on_document(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"order 12345 shipped"'),
        )
        doc = Document(document_elm, Mock())

        matches = doc.search_regex(r"\d+")

        assert len(matches) == 1
        assert matches[0].start == 6
        assert matches[0].end == 11

    def it_exposes_replace_regex_on_document(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"foobar"'),
        )
        doc = Document(document_elm, Mock())

        count = doc.replace_regex(r"(foo)bar", r"\1baz")

        assert count == 1
        assert doc.paragraphs[0].text == "foobaz"

    def it_passes_flags_through_to_search_regex(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"Hello HELLO"'),
        )
        doc = Document(document_elm, Mock())

        matches = doc.search_regex(r"hello", re.IGNORECASE)

        assert len(matches) == 2

    def it_passes_flags_through_to_replace_regex(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"Hello HELLO"'),
        )
        doc = Document(document_elm, Mock())

        count = doc.replace_regex(r"hello", "hi", re.IGNORECASE)

        assert count == 2
        assert doc.paragraphs[0].text == "hi hi"

    def it_accepts_a_compiled_pattern_on_search_regex(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"Hello HELLO"'),
        )
        doc = Document(document_elm, Mock())
        compiled = re.compile(r"hello", re.IGNORECASE)

        matches = doc.search_regex(compiled)

        assert len(matches) == 2

    def it_accepts_a_compiled_pattern_on_replace_regex(self):
        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"Hello HELLO"'),
        )
        doc = Document(document_elm, Mock())
        compiled = re.compile(r"hello", re.IGNORECASE)

        count = doc.replace_regex(compiled, "hi")

        assert count == 2
        assert doc.paragraphs[0].text == "hi hi"
