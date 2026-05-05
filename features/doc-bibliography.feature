Feature: Read and write the bibliography / citation sources
  In order to author and introspect citation sources in a document
  As a developer using python-docx
  I need a Document.bibliography collection plus Document.add_citation
  and Paragraph.add_citation_reference


  Scenario: Document.bibliography starts empty on a fresh document
    Given a fresh default document
     Then document.bibliography has length 0


  Scenario: Document.add_citation appends a source reachable via bibliography
    Given a fresh default document
     When I call document.add_citation("smith2020", title="A Book", author="Smith, J.", year=2020)
     Then document.bibliography has length 1
      And document.bibliography.get_by_tag("smith2020").title is "A Book"
      And document.bibliography.get_by_tag("smith2020").year is "2020"


  Scenario: Paragraph.add_citation_reference emits a citation SDT
    Given a fresh default document
     When I call document.add_citation("einstein1905", title="Relativity", author="Einstein, A.", year=1905)
      And I add a paragraph with a citation reference for "einstein1905"
     Then the last paragraph contains a citation sdt referencing "einstein1905"


  Scenario: Bibliography survives a save-reload roundtrip
    Given a fresh default document
     When I call document.add_citation("keynes1936", title="The General Theory", author="Keynes, J.M.", year=1936)
      And I save and reload the document
     Then document.bibliography has length 1
      And document.bibliography.get_by_tag("keynes1936").title is "The General Theory"
