Feature: Mutate footnotes in a document
  In order to edit or remove footnotes already present in a document
  As a developer using python-docx
  I need mutation methods on Footnote objects


  Scenario: Footnote.delete() removes the footnote from the part
    Given a document having 3 footnotes
     When I delete the footnote with id 3
     Then len(document.footnotes) == 2
      And the yielded footnote ids are [2, 4]


  Scenario: Footnote.delete() removes the footnoteReference from the body
    Given a document having 3 footnotes
     When I delete the footnote with id 3
     Then no footnoteReference with id 3 remains in the document body


  Scenario: Footnote.clear() leaves a single empty paragraph
    Given a document having 3 footnotes
     When I clear the footnote with id 2
     Then footnote with id 2 has text ""
      And footnote with id 2 has 1 paragraph
      And footnote with id 2 paragraph has the FootnoteText style


  Scenario: Footnote.add_paragraph() appends a paragraph with FootnoteText style
    Given a document having 3 footnotes
     When I call footnote.add_paragraph("second line") on footnote with id 2
     Then footnote with id 2 has 2 paragraphs
      And the new paragraph has text "second line"
      And the new paragraph has the FootnoteText style


  Scenario: Deletion survives a save/open round-trip
    Given a document having 3 footnotes
     When I delete the footnote with id 3
      And I save and reopen the document
     Then len(document.footnotes) == 2
      And the yielded footnote ids are [2, 4]
