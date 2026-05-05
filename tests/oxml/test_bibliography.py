# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.bibliography` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.bibliography import CT_Source, CT_Sources, new_sources_root
from docx.oxml.ns import qn


def _empty_sources() -> CT_Sources:
    return new_sources_root()


class DescribeCT_Sources:
    """Unit-test suite for `docx.oxml.bibliography.CT_Sources`."""

    def it_starts_with_an_empty_source_lst(self):
        sources = _empty_sources()

        assert sources.source_lst == []

    def it_carries_default_selected_style_and_style_name(self):
        sources = _empty_sources()

        assert sources.selected_style == "/APA.XSL"
        assert sources.style_name == "APA"

    def it_can_add_a_simple_source(self):
        sources = _empty_sources()

        src = sources.add_source_from_kwargs(
            "smith2020", title="Test Book", author="Smith, J.", year=2020
        )

        assert isinstance(src, CT_Source)
        assert src.tag_val == "smith2020"
        assert src.title == "Test Book"
        assert src.year == "2020"
        assert src.source_type == "Book"
        assert src.author == "Smith, J."

    def it_defaults_source_type_to_Book(self):
        sources = _empty_sources()

        src = sources.add_source_from_kwargs("x")

        assert src.source_type == "Book"

    def it_respects_explicit_source_type(self):
        sources = _empty_sources()

        src = sources.add_source_from_kwargs("x", source_type="JournalArticle")

        assert src.source_type == "JournalArticle"

    def it_exposes_extra_kwargs_as_text_children(self):
        sources = _empty_sources()

        src = sources.add_source_from_kwargs("x", city="London", publisher="Acme")

        # -- two direct children with capitalized tag names --
        city = src.find(qn("b:City"))
        publisher = src.find(qn("b:Publisher"))
        assert city is not None and city.text == "London"
        assert publisher is not None and publisher.text == "Acme"

    def it_can_look_up_a_source_by_tag(self):
        sources = _empty_sources()
        sources.add_source_from_kwargs("alpha")
        target = sources.add_source_from_kwargs("beta")
        sources.add_source_from_kwargs("gamma")

        assert sources.get_source_by_tag("beta") is target
        assert sources.get_source_by_tag("missing") is None

    def it_allows_clearing_the_selected_style(self):
        sources = _empty_sources()

        sources.selected_style = None
        sources.style_name = None

        assert sources.selected_style is None
        assert sources.style_name is None

    def it_appends_each_new_source_to_the_end(self):
        sources = _empty_sources()

        sources.add_source_from_kwargs("a")
        sources.add_source_from_kwargs("b")
        sources.add_source_from_kwargs("c")

        tags = [s.tag_val for s in sources.source_lst]
        assert tags == ["a", "b", "c"]


class DescribeCT_Source:
    """Unit-test suite for `docx.oxml.bibliography.CT_Source`."""

    def its_author_falls_back_to_the_person_NameList_form(self):
        sources = _empty_sources()
        # -- build a source with a Person-style Author block by hand --
        from docx.oxml.parser import OxmlElement

        src = sources.add_source_from_kwargs("k")
        # -- remove the Corporate-style Author the helper generated --
        for author_root in src.findall(qn("b:Author")):
            src.remove(author_root)
        b_ns = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"

        def _e(tag: str, text: "str | None" = None):
            e = OxmlElement(f"b:{tag}", nsdecls={"b": b_ns})
            if text is not None:
                e.text = text
            return e

        author_root = _e("Author")
        inner = _e("Author")
        name_list = _e("NameList")
        person = _e("Person")
        person.append(_e("First", "Jane"))
        person.append(_e("Last", "Doe"))
        name_list.append(person)
        inner.append(name_list)
        author_root.append(inner)
        src.append(author_root)

        assert src.author == "Jane Doe"

    def its_author_is_None_when_no_author_is_set(self):
        sources = _empty_sources()
        src = sources.add_source_from_kwargs("k")
        # -- drop the helper-added Author --
        for author_root in src.findall(qn("b:Author")):
            src.remove(author_root)

        assert src.author is None
