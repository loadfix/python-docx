Feature: Rendered list-label for numbered paragraphs
  In order to display or export the number/bullet Word would show next to a list item
  As a developer using python-docx
  I need Paragraph.list_label and Document.list_labels()


  Scenario: list_label is None for a paragraph that isn't in a list
    Given a fresh default document with one paragraph
     Then the paragraph's list_label is None


  Scenario: Decimal list labels advance across paragraphs
    Given a document with a decimal-then-letter multi-level numbering and 6 nested paragraphs
     Then the list labels for the 6 paragraphs are "1., a), b), 2., 3., a)"


  Scenario: Bullet level renders its lvlText verbatim
    Given a document with a single-level bullet numbering and 3 bullet paragraphs
     Then the list labels for the 3 paragraphs are "•, •, •"


  Scenario: Document.list_labels returns a mapping for every numbered paragraph
    Given a document with a decimal-then-letter multi-level numbering and 6 nested paragraphs
     Then document.list_labels has an entry for each of the 6 paragraphs
      And the list_labels entry for paragraph 1 is "1."
      And the list_labels entry for paragraph 4 is "2."


  Scenario: A paragraph not in a list is absent from document.list_labels
    Given a document with a decimal-then-letter multi-level numbering and 6 nested paragraphs
      And a trailing plain paragraph is appended
     Then document.list_labels has no entry for the trailing paragraph
