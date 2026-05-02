Feature: Generate a table of contents
  In order to add a navigable heading index to a Word document
  As a developer using python-docx
  I need a way to write a TOC field that lists the document's headings


  Scenario: Document.add_table_of_contents() using defaults
    Given a document having heading paragraphs
     When I assign toc = document.add_table_of_contents()
     Then toc is a Paragraph object
      And toc is the last paragraph in the document
      And toc has one complex TOC field
      And the TOC field instruction is ' TOC \o "1-3" \h \z \u '
      And the TOC preview lists 6 entries


  Scenario: Document.add_table_of_contents(levels=(1, 1)) limits to H1
    Given a document having heading paragraphs
     When I assign toc = document.add_table_of_contents(levels=(1, 1))
     Then the TOC field instruction is ' TOC \o "1-1" \h \z \u '
      And the TOC preview lists 2 entries
      And the TOC preview contains "Chapter One"
      And the TOC preview does not contain "Section 1.1"


  Scenario: Document.add_table_of_contents(levels=(2, 3)) skips H1
    Given a document having heading paragraphs
     When I assign toc = document.add_table_of_contents(levels=(2, 3))
     Then the TOC field instruction is ' TOC \o "2-3" \h \z \u '
      And the TOC preview lists 4 entries
      And the TOC preview does not contain "Chapter One"
      And the TOC preview contains "Section 1.1"
      And the TOC preview contains "Subsection 1.1.1"


  Scenario: Paragraph.insert_table_of_contents_before() inserts at position
    Given a document having heading paragraphs
     When I assign toc = anchor.insert_table_of_contents_before()
     Then toc is a Paragraph object
      And toc is the first paragraph in the document
      And toc has one complex TOC field
      And the TOC preview lists 6 entries


  Scenario: Paragraph.insert_table_of_contents_after() inserts at position
    Given a document having heading paragraphs
     When I assign toc = anchor.insert_table_of_contents_after()
     Then toc is a Paragraph object
      And toc is the paragraph after the anchor
      And toc has one complex TOC field
      And the TOC preview lists 6 entries


  Scenario: Document.add_table_of_contents() on an empty document
    Given a document having no heading paragraphs
     When I assign toc = document.add_table_of_contents()
     Then toc is a Paragraph object
      And toc has one complex TOC field
      And the TOC field instruction is ' TOC \o "1-3" \h \z \u '
      And the TOC preview lists 0 entries


  Scenario Outline: Document.add_table_of_contents() rejects invalid levels
    Given a document having heading paragraphs
     When I call document.add_table_of_contents(levels=<levels>) expecting ValueError
     Then a ValueError is raised

    Examples: invalid level tuples
      | levels  |
      | (0, 3)  |
      | (1, 10) |
      | (3, 1)  |
