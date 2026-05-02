Feature: Add a DrawingML preset shape to a paragraph
  In order to author documents containing native Word shapes
  As a developer using python-docx
  I need a way to add a preset shape inline to a paragraph


  Scenario Outline: Add a preset shape to a paragraph
    Given a paragraph
     When I add a preset shape of type <shape_type> to the paragraph
     Then the paragraph has 1 inline drawing
      And the drawing type is SHAPE
      And the wps:wsp shape type is <shape_type>

    Examples: preset shapes
      | shape_type                |
      | RECTANGLE                 |
      | ROUNDED_RECTANGLE         |
      | OVAL                      |
      | ARROW_RIGHT               |
      | CALLOUT_ROUNDED_RECTANGLE |


  Scenario: Add a preset shape carrying text
    Given a paragraph
     When I add a preset shape of type RECTANGLE with text "Hello"
     Then the paragraph has 1 inline drawing
      And the drawing type is TEXT_BOX
      And the wps:wsp shape text is "Hello"


  Scenario: Reject a non-WD_SHAPE argument
    Given a paragraph
     Then calling add_shape with a non-WD_SHAPE argument raises TypeError


  Scenario: Read preset shapes from a fixture
    Given a document known to contain five inline preset shapes
     Then the document's third inline drawing is a SHAPE of type OVAL
      And the document's fifth inline drawing is a TEXT_BOX with text "Hello shape"
