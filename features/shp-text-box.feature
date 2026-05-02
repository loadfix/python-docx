Feature: Read text content of a DrawingML text box
  In order to extract copy written inside shapes
  As a developer using python-docx
  I need a way to access the text frame contents of a ``wps:wsp`` shape


  Scenario: Drawing with a single-paragraph text frame
    Given a document known to contain shape text frames
     Then the first shape text frame has text "Single line"
      And the first shape text frame has 1 paragraph


  Scenario: Drawing with a multi-paragraph text frame
    Given a document known to contain shape text frames
     Then the second shape text frame has text spanning 3 paragraphs
      And the second shape text frame exposes the expected paragraph texts


  Scenario: Replace text in a newly-created shape
    Given a paragraph
     When I add a preset shape of type RECTANGLE with text "Initial"
     Then the wps:wsp shape text is "Initial"
      When I set the shape's text to "Replaced"
      Then the wps:wsp shape text is "Replaced"


  Scenario: Empty shape returns empty text
    Given a paragraph
     When I add a preset shape of type RECTANGLE to the paragraph
     Then the wps:wsp shape text is ""
