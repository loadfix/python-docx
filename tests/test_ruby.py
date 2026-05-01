"""Unit-test suite for `docx.ruby` and its oxml counterparts."""

from __future__ import annotations

from typing import cast

from docx.oxml.ruby import CT_Ruby
from docx.ruby import RubyAnnotation

from .unitutil.cxml import element


class DescribeRubyAnnotation:
    """Unit-test suite for `docx.ruby.RubyAnnotation`."""

    def it_reads_base_text_and_ruby_text(self):
        ruby = cast(
            CT_Ruby,
            element(
                'w:ruby/(w:rubyPr,'
                'w:rt/w:r/w:t"nihon",'
                'w:rubyBase/w:r/w:t"Japan")'
            ),
        )
        ann = RubyAnnotation(ruby)
        assert ann.base_text == "Japan"
        assert ann.ruby_text == "nihon"

    def it_returns_empty_strings_when_components_absent(self):
        ruby = cast(CT_Ruby, element("w:ruby"))
        ann = RubyAnnotation(ruby)
        assert ann.base_text == ""
        assert ann.ruby_text == ""

    def it_reads_alignment(self):
        ruby = cast(
            CT_Ruby,
            element("w:ruby/w:rubyPr/w:rubyAlign{w:val=distributeSpace}"),
        )
        ann = RubyAnnotation(ruby)
        assert ann.alignment == "distributeSpace"

    def it_returns_None_for_alignment_when_absent(self):
        ruby = cast(CT_Ruby, element("w:ruby/w:rubyPr"))
        ann = RubyAnnotation(ruby)
        assert ann.alignment is None

    def it_returns_None_for_alignment_when_no_rubyPr(self):
        ruby = cast(CT_Ruby, element("w:ruby"))
        ann = RubyAnnotation(ruby)
        assert ann.alignment is None

    def it_reads_language(self):
        ruby = cast(
            CT_Ruby,
            element("w:ruby/w:rubyPr/w:lid{w:val=ja-JP}"),
        )
        ann = RubyAnnotation(ruby)
        assert ann.language == "ja-JP"

    def it_returns_None_for_language_when_absent(self):
        ruby = cast(CT_Ruby, element("w:ruby/w:rubyPr"))
        ann = RubyAnnotation(ruby)
        assert ann.language is None


class DescribeCT_Ruby_text:
    """Confirm `w:ruby` base text is included in run.text extraction."""

    def it_contributes_base_text_to_run_text(self):
        r = element(
            'w:r/(w:t"before ",'
            'w:ruby/(w:rubyPr,w:rt/w:r/w:t"rt",w:rubyBase/w:r/w:t"BASE"),'
            'w:t" after")'
        )
        # CT_R.text concatenates via xpath that now includes w:ruby
        assert r.text == "before BASE after"

    def it_yields_multiple_rubies(self):
        r = element(
            'w:r/('
            'w:ruby/(w:rubyPr,w:rt/w:r/w:t"a1",w:rubyBase/w:r/w:t"A"),'
            'w:ruby/(w:rubyPr,w:rt/w:r/w:t"b1",w:rubyBase/w:r/w:t"B"))'
        )
        rubies = r.ruby_lst
        assert len(rubies) == 2
        assert rubies[0].base_text == "A"
        assert rubies[1].base_text == "B"
