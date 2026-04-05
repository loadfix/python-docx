Feature: Access and modify footnotes
  In order to work with footnotes in a Word document
  As a developer using python-docx
  I need methods to access and mutate footnotes


  Scenario: Access footnotes collection from a new document
    Given a new document
     Then document.footnotes is a Footnotes object
      And len(document.footnotes) == 0


  Scenario: Add a footnote to a document
    Given a new document with a paragraph
     When I add a footnote to the first run
     Then len(document.footnotes) == 1
      And the footnote has a single paragraph with FootnoteText style


  Scenario: Add a footnote with text
    Given a new document with a paragraph
     When I add a footnote with text "See reference." to the first run
     Then the footnote text is "See reference."


  Scenario: Iterate over footnotes
    Given a new document with two footnotes
     Then iterating document.footnotes yields 2 Footnote objects


  Scenario: Clear a footnote
    Given a new document with a footnote containing text
     When I clear the footnote
     Then the footnote has a single paragraph with FootnoteText style
      And the footnote text is empty


  Scenario: Delete a footnote
    Given a new document with a footnote
     When I delete the footnote
     Then len(document.footnotes) == 0
