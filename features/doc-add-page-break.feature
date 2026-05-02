Feature: Add a page break
  In order to force a page break at a particular location
  As a developer using the python-docx
  I need a way to add a hard page break on its own paragraph


  Scenario: Add a hard page break paragraph
    Given a blank document
     When I add a page break to the document
     Then the last paragraph contains only a page break


  Scenario: Remove a page break from a paragraph via Paragraph.clear_page_breaks
    Given a blank document
     When I add a page break to the document
      And I clear page breaks on the last paragraph
     Then the last paragraph has no page break


  Scenario: Paragraph.clear_page_breaks is a no-op on a paragraph with no page break
    Given a document containing five paragraphs
     When I clear page breaks on the third paragraph
     Then the third paragraph has no page break
