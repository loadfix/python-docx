Feature: Add a footnote to a document
  In order to add a footnote to a document
  As a developer using python-docx
  I need a way to add a footnote anchored on a specific run


  Scenario: Footnotes.add(run) with no text
    Given a blank document
     When I add a paragraph with text "Body prose."
      And I assign footnote = document.footnotes.add(paragraph.runs[0])
     Then footnote is a Footnote object
      And footnote.footnote_id == 2
      And len(footnote.paragraphs) == 1
      And footnote.paragraphs[0].style.name == "FootnoteText"
      And len(document.footnotes) == 1


  Scenario: Footnotes.add(run, text) with footnote text
    Given a blank document
     When I add a paragraph with text "Body prose."
      And I assign footnote = document.footnotes.add(paragraph.runs[0], "See note")
     Then footnote.text == "See note"
      And footnote.footnote_id == 2


  Scenario: Footnotes.add() inserts a footnoteReference into the anchor run
    Given a blank document
     When I add a paragraph with text "Body prose."
      And I assign footnote = document.footnotes.add(paragraph.runs[0], "See note")
     Then the anchor run contains a footnoteReference to footnote.footnote_id
      And the anchor run has the FootnoteReference character style


  Scenario: Footnotes.add() assigns sequential ids starting at 2
    Given a blank document
     When I add a paragraph with text "First."
      And I assign fn1 = document.footnotes.add(paragraph.runs[0], "one")
      And I add a paragraph with text "Second."
      And I assign fn2 = document.footnotes.add(paragraph.runs[0], "two")
     Then fn1.footnote_id == 2
      And fn2.footnote_id == 3
      And len(document.footnotes) == 2


  Scenario: Added footnote survives a save/open round-trip
    Given a blank document
     When I add a paragraph with text "Body prose."
      And I assign footnote = document.footnotes.add(paragraph.runs[0], "See note")
      And I save and reopen the document
     Then len(document.footnotes) == 1
      And document.footnotes[0].text == "See note"
      And document.footnotes[0].footnote_id == 2
