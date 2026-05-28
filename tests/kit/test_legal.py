"""Unit-test suite for ``docx.kit.legal`` template factories."""

from __future__ import annotations

import pytest

from docx.document import Document as DocumentCls
from docx.kit import legal


# -- Shared helpers -------------------------------------------------------


def _texts(document: DocumentCls):
    """Return the text of every paragraph in ``document``."""
    return [p.text for p in document.paragraphs]


def _full_text(document: DocumentCls) -> str:
    return "\n".join(_texts(document))


@pytest.fixture
def parties():
    return [
        {"role": "Plaintiff", "name": "Acme Corp Pty Ltd"},
        {"role": "Defendant", "name": "Beta Pty Ltd"},
    ]


@pytest.fixture
def body():
    return [
        {
            "heading": "Background",
            "paragraphs": [
                "The Plaintiff is incorporated in NSW.",
                "The Defendant operates a competing business.",
            ],
        },
        {
            "heading": "Cause of Action",
            "paragraphs": ["The Defendant has breached the contract."],
        },
    ]


# -- court_paper ----------------------------------------------------------


class DescribeCourtPaper:
    """Unit-test suite for ``legal.court_paper``."""

    def it_returns_a_document_with_the_court_heading(self, parties):
        doc = legal.court_paper(parties=parties)

        assert isinstance(doc, DocumentCls)
        text = _full_text(doc)
        assert "FEDERAL COURT OF AUSTRALIA" in text

    def it_renders_the_supplied_division(self, parties):
        doc = legal.court_paper(
            parties=parties, division="Victoria District Registry"
        )

        assert "Victoria District Registry" in _full_text(doc)

    def it_renders_the_case_number(self, parties):
        doc = legal.court_paper(
            parties=parties, case_no="NSD 1234 of 2026"
        )

        assert "No. NSD 1234 of 2026" in _full_text(doc)

    def it_renders_the_document_type_subtitle(self, parties):
        doc = legal.court_paper(
            parties=parties, document_type="Statement of Claim"
        )

        assert "Statement of Claim" in _full_text(doc)

    def it_renders_the_parties_block(self, parties):
        doc = legal.court_paper(parties=parties)
        text = _full_text(doc)

        assert "BETWEEN:" in text
        assert "Acme Corp Pty Ltd" in text
        assert "Beta Pty Ltd" in text
        assert "Plaintiff" in text
        assert "Defendant" in text

    def it_includes_the_disclaimer(self, parties):
        doc = legal.court_paper(parties=parties)
        text = _full_text(doc)

        assert "DISCLAIMER" in text
        assert "not legal advice" in text

    def it_renders_numbered_body_paragraphs(self, parties, body):
        doc = legal.court_paper(parties=parties, body=body)
        text = _full_text(doc)

        # -- three body paragraphs across two sections --
        assert "1. The Plaintiff is incorporated in NSW." in text
        assert "2. The Defendant operates a competing business." in text
        assert "3. The Defendant has breached the contract." in text

    def it_renders_the_section_headings(self, parties, body):
        doc = legal.court_paper(parties=parties, body=body)
        text = _full_text(doc)

        assert "Background" in text
        assert "Cause of Action" in text

    def it_enables_line_numbering_when_requested(self, parties):
        doc = legal.court_paper(parties=parties, line_numbering=True)

        section = doc.sections[0]
        assert section.line_numbering is not None
        assert section.line_numbering.count_by == 1

    def it_does_not_enable_line_numbering_by_default(self, parties):
        doc = legal.court_paper(parties=parties)

        assert doc.sections[0].line_numbering is None

    def it_accepts_a_dict_for_line_numbering(self, parties):
        doc = legal.court_paper(
            parties=parties,
            line_numbering={"count_by": 5, "start": 1},
        )

        ln = doc.sections[0].line_numbering
        assert ln is not None
        assert ln.count_by == 5
        assert ln.start == 1

    def it_includes_a_signature_block(self, parties):
        doc = legal.court_paper(parties=parties)
        text = _full_text(doc)

        assert "Filed by" in text
        assert "Signed for and on behalf of Acme Corp Pty Ltd" in text
        assert "Name: ______________________________" in text

    def it_raises_when_parties_missing(self):
        with pytest.raises(ValueError, match="parties is required"):
            legal.court_paper()

    def it_raises_when_parties_empty(self):
        with pytest.raises(ValueError, match="at least one party"):
            legal.court_paper(parties=[])

    def it_raises_when_a_party_lacks_a_name(self):
        with pytest.raises(ValueError, match="non-empty 'name'"):
            legal.court_paper(parties=[{"role": "Plaintiff"}])


# -- table_of_authorities -------------------------------------------------


class DescribeTableOfAuthorities:
    """Unit-test suite for ``legal.table_of_authorities``."""

    def it_returns_a_document_with_the_default_title(self):
        doc = legal.table_of_authorities()

        assert isinstance(doc, DocumentCls)
        assert "Table of Authorities" in _full_text(doc)

    def it_renders_a_custom_title(self):
        doc = legal.table_of_authorities(title="Cases Cited")

        assert "Cases Cited" in _full_text(doc)

    def it_renders_the_supplied_citations(self):
        doc = legal.table_of_authorities(
            citations=[
                {"case": "Donoghue v Stevenson [1932] AC 562", "first_pin": 580},
                {"case": "Smith v Jones (2020) 270 CLR 100", "first_pin": 105},
            ],
        )
        text = _full_text(doc)

        assert "1. Donoghue v Stevenson [1932] AC 562 at 580" in text
        assert "2. Smith v Jones (2020) 270 CLR 100 at 105" in text

    def it_renders_a_citation_without_a_pinpoint(self):
        doc = legal.table_of_authorities(
            citations=[{"case": "Mabo v Queensland (No 2) (1992) 175 CLR 1"}]
        )

        text = _full_text(doc)
        assert "1. Mabo v Queensland (No 2) (1992) 175 CLR 1" in text
        # -- no "at" appended when first_pin is omitted --
        assert "Mabo v Queensland (No 2) (1992) 175 CLR 1 at" not in text

    def it_falls_back_to_a_placeholder_when_no_citations_supplied(self):
        doc = legal.table_of_authorities()

        assert "[Insert citations here.]" in _full_text(doc)

    def it_emits_a_TOA_field(self):
        doc = legal.table_of_authorities(
            citations=[{"case": "Donoghue v Stevenson [1932] AC 562"}]
        )

        # -- the TOA field surfaces via the document's fields collection --
        toa_fields = [
            f
            for f in doc.fields
            if f.field_type and f.field_type.upper() == "TOA"
        ]
        assert len(toa_fields) == 1

    def it_emits_a_categorised_TOA_field_when_category_supplied(self):
        doc = legal.table_of_authorities(
            citations=[{"case": "Donoghue v Stevenson [1932] AC 562"}],
            category=1,
        )

        toa_fields = [
            f
            for f in doc.fields
            if f.field_type and f.field_type.upper() == "TOA"
        ]
        assert len(toa_fields) == 1
        # -- the \c "1" switch should appear in the field instruction --
        assert '\\c "1"' in toa_fields[0].instruction

    def it_includes_the_disclaimer(self):
        doc = legal.table_of_authorities(
            citations=[{"case": "Donoghue v Stevenson [1932] AC 562"}]
        )

        assert "DISCLAIMER" in _full_text(doc)

    def it_raises_when_a_citation_lacks_a_case_key(self):
        with pytest.raises(ValueError, match="non-empty 'case'"):
            legal.table_of_authorities(citations=[{"first_pin": 105}])

    def it_raises_when_a_citation_is_not_a_mapping(self):
        with pytest.raises(ValueError, match="must be a mapping"):
            legal.table_of_authorities(citations=["plain string"])


# -- brief ----------------------------------------------------------------


class DescribeBrief:
    """Unit-test suite for ``legal.brief``."""

    def it_builds_a_brief_with_a_matter_title(self):
        doc = legal.brief(
            matter="Smith v Jones",
            counsel={"name": "Bob Wilson SC"},
        )
        text = _full_text(doc)

        assert "BRIEF TO COUNSEL" in text
        assert "Smith v Jones" in text

    def it_renders_the_counsel_block(self):
        doc = legal.brief(
            matter="Smith v Jones",
            counsel={
                "name": "Bob Wilson SC",
                "chambers": "5 Wentworth Chambers",
                "email": "bob@5wentworth.com.au",
            },
        )
        text = _full_text(doc)

        assert "Bob Wilson SC" in text
        assert "5 Wentworth Chambers" in text
        assert "bob@5wentworth.com.au" in text

    def it_renders_the_instructing_solicitor_block(self):
        doc = legal.brief(
            matter="Smith v Jones",
            counsel={"name": "Bob Wilson SC"},
            instructing_solicitor={
                "name": "Alice Smith",
                "firm": "Smith & Associates",
                "email": "alice@smith.com.au",
            },
        )
        text = _full_text(doc)

        assert "Instructing Solicitor" in text
        assert "Alice Smith" in text
        assert "Smith & Associates" in text

    def it_omits_the_solicitor_block_when_not_supplied(self):
        doc = legal.brief(
            matter="Smith v Jones",
            counsel={"name": "Bob Wilson SC"},
        )

        # -- the heading-style "Instructing Solicitor" should not appear
        # -- as a section heading. The signature-block label "Brief
        # -- delivered by ... Instructing Solicitor" is fine. --
        text = _full_text(doc)
        # -- counsel block is emitted, instructing-solicitor heading is not --
        assert "Counsel" in text
        # -- no firm field is rendered --
        assert "Smith & Associates" not in text

    def it_renders_the_court_heading_when_parties_supplied(self, parties):
        doc = legal.brief(
            matter="Smith v Jones",
            counsel={"name": "Bob Wilson SC"},
            parties=parties,
            case_no="NSD 1234 of 2026",
        )
        text = _full_text(doc)

        assert "FEDERAL COURT OF AUSTRALIA" in text
        assert "BETWEEN:" in text

    def it_renders_numbered_observation_paragraphs(self):
        doc = legal.brief(
            matter="Smith v Jones",
            counsel={"name": "Bob Wilson SC"},
            sections=[
                {
                    "heading": "Observations",
                    "paragraphs": ["First observation.", "Second observation."],
                }
            ],
        )
        text = _full_text(doc)

        assert "1. First observation." in text
        assert "2. Second observation." in text

    def it_renders_the_documents_index(self):
        doc = legal.brief(
            matter="Smith v Jones",
            counsel={"name": "Bob Wilson SC"},
            documents_index=[
                "Statement of Claim",
                "Defence",
                "Witness statement of Alice Smith",
            ],
        )
        text = _full_text(doc)

        assert "Index of Documents" in text
        assert "1. Statement of Claim" in text
        assert "3. Witness statement of Alice Smith" in text

    def it_enables_line_numbering_when_requested(self):
        doc = legal.brief(
            matter="Smith v Jones",
            counsel={"name": "Bob Wilson SC"},
            line_numbering=True,
        )

        assert doc.sections[0].line_numbering is not None

    def it_includes_the_disclaimer(self):
        doc = legal.brief(
            matter="Smith v Jones",
            counsel={"name": "Bob Wilson SC"},
        )

        assert "DISCLAIMER" in _full_text(doc)

    def it_raises_when_matter_is_missing(self):
        with pytest.raises(ValueError, match="matter is required"):
            legal.brief(counsel={"name": "Bob"})

    def it_raises_when_counsel_is_missing(self):
        with pytest.raises(
            ValueError, match="counsel must be a mapping with a non-empty 'name'"
        ):
            legal.brief(matter="Smith v Jones")

    def it_raises_when_counsel_lacks_a_name(self):
        with pytest.raises(
            ValueError, match="counsel must be a mapping with a non-empty 'name'"
        ):
            legal.brief(matter="Smith v Jones", counsel={"chambers": "X"})


# -- declaration ----------------------------------------------------------


class DescribeDeclaration:
    """Unit-test suite for ``legal.declaration``."""

    def it_builds_a_declaration_with_the_declarant_name(self):
        doc = legal.declaration(
            declarant={"name": "Alice Smith", "address": "1 Smith St, Sydney"},
        )
        text = _full_text(doc)

        assert "DECLARATION" in text
        assert "I, Alice Smith" in text
        assert "1 Smith St, Sydney" in text

    def it_renders_the_occupation_when_supplied(self):
        doc = legal.declaration(
            declarant={
                "name": "Alice Smith",
                "occupation": "solicitor",
                "address": "1 Smith St",
            },
        )
        text = _full_text(doc)

        assert "I, Alice Smith, solicitor, of 1 Smith St" in text

    def it_renders_numbered_statements_of_fact(self):
        doc = legal.declaration(
            declarant={"name": "Alice Smith"},
            paragraphs=[
                "I am the Plaintiff in this proceeding.",
                "I have personal knowledge of the matters stated.",
            ],
        )
        text = _full_text(doc)

        assert "1. I am the Plaintiff in this proceeding." in text
        assert "2. I have personal knowledge of the matters stated." in text

    def it_renders_the_jurat(self):
        doc = legal.declaration(
            declarant={"name": "Alice Smith"},
            affirmed_at="Sydney",
            affirmed_date="2026-03-15",
        )
        text = _full_text(doc)

        assert "Affirmed at Sydney on 2026-03-15." in text

    def it_renders_a_jurat_placeholder_when_date_missing(self):
        doc = legal.declaration(declarant={"name": "Alice Smith"})

        text = _full_text(doc)
        assert "Affirmed at Sydney on" in text

    def it_renders_signature_stubs_for_declarant_and_witness(self):
        doc = legal.declaration(declarant={"name": "Alice Smith"})

        text = _full_text(doc)
        assert "Declarant:" in text
        assert "(Alice Smith)" in text
        assert "Witness:" in text
        assert "Witness Qualification" in text

    def it_renders_the_court_heading_when_parties_supplied(self, parties):
        doc = legal.declaration(
            declarant={"name": "Alice Smith"},
            parties=parties,
            case_no="NSD 1234 of 2026",
        )
        text = _full_text(doc)

        assert "FEDERAL COURT OF AUSTRALIA" in text
        assert "Acme Corp Pty Ltd" in text

    def it_renders_a_placeholder_when_no_paragraphs_supplied(self):
        doc = legal.declaration(declarant={"name": "Alice Smith"})

        assert "[Insert numbered statements of fact here.]" in _full_text(doc)

    def it_enables_line_numbering_when_requested(self):
        doc = legal.declaration(
            declarant={"name": "Alice Smith"}, line_numbering=True
        )

        assert doc.sections[0].line_numbering is not None

    def it_includes_the_disclaimer(self):
        doc = legal.declaration(declarant={"name": "Alice Smith"})

        assert "DISCLAIMER" in _full_text(doc)

    def it_raises_when_declarant_is_missing(self):
        with pytest.raises(
            ValueError, match="declarant must be a mapping with a non-empty 'name'"
        ):
            legal.declaration()

    def it_raises_when_declarant_lacks_a_name(self):
        with pytest.raises(
            ValueError, match="declarant must be a mapping with a non-empty 'name'"
        ):
            legal.declaration(declarant={"address": "1 Smith St"})


# -- Round-trip integration ----------------------------------------------


class DescribeLegalRoundTrip:
    """End-to-end smoke-tests: every factory produces a saveable document."""

    def it_can_save_a_court_paper_to_a_BytesIO(self, parties, body):
        from io import BytesIO

        doc = legal.court_paper(
            court="Federal Court of Australia",
            division="New South Wales District Registry",
            case_no="NSD 1234 of 2026",
            parties=parties,
            document_type="Statement of Claim",
            line_numbering=True,
            body=body,
        )
        buf = BytesIO()
        doc.save(buf)
        # -- Word .docx is a zip; the magic bytes are 'PK' --
        assert buf.getvalue()[:2] == b"PK"

    def it_can_save_a_table_of_authorities_to_a_BytesIO(self):
        from io import BytesIO

        doc = legal.table_of_authorities(
            citations=[
                {"case": "Donoghue v Stevenson [1932] AC 562", "first_pin": 580},
                {"case": "Smith v Jones (2020) 270 CLR 100", "first_pin": 105},
            ],
        )
        buf = BytesIO()
        doc.save(buf)
        assert buf.getvalue()[:2] == b"PK"

    def it_can_save_a_brief_to_a_BytesIO(self):
        from io import BytesIO

        doc = legal.brief(
            matter="Smith v Jones",
            counsel={"name": "Bob Wilson SC", "chambers": "5 Wentworth"},
            instructing_solicitor={"name": "Alice", "firm": "Smith & Co"},
            sections=[{"heading": "Observations", "paragraphs": ["..."]}],
            documents_index=["Statement of Claim", "Defence"],
        )
        buf = BytesIO()
        doc.save(buf)
        assert buf.getvalue()[:2] == b"PK"

    def it_can_save_a_declaration_to_a_BytesIO(self):
        from io import BytesIO

        doc = legal.declaration(
            declarant={"name": "Alice Smith", "address": "1 Smith St"},
            affirmed_at="Sydney",
            affirmed_date="2026-03-15",
            paragraphs=["I am the Plaintiff.", "I know the facts."],
        )
        buf = BytesIO()
        doc.save(buf)
        assert buf.getvalue()[:2] == b"PK"
