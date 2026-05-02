Feature: Add and read floating (anchored) pictures
  In order to place pictures outside the flow of text
  As a developer using python-docx
  I need a way to add, configure, and inspect floating images


  Scenario: Add a floating image with default position
    Given a paragraph
     When I add a floating image to the paragraph with no position
     Then the paragraph has 1 floating image
      And the floating image has horizontal anchor COLUMN
      And the floating image has vertical anchor PARAGRAPH
      And the floating image has wrap type SQUARE


  Scenario Outline: Add a floating image with explicit position
    Given a paragraph
     When I add a floating image anchored <h_anchor>/<v_anchor> with wrap <wrap>
     Then the floating image has horizontal anchor <h_anchor>
      And the floating image has vertical anchor <v_anchor>
      And the floating image has wrap type <wrap>

    Examples: positioning combinations
      | h_anchor  | v_anchor  | wrap          |
      | PAGE      | PAGE      | TIGHT         |
      | MARGIN    | MARGIN    | TOP_AND_BOTTOM|
      | COLUMN    | PARAGRAPH | BEHIND        |
      | PAGE      | LINE      | IN_FRONT      |


  Scenario: Read floating images from a fixture
    Given a document known to contain three floating images
     Then the document has 3 floating images across its paragraphs
      And the second floating image has horizontal anchor PAGE
      And the second floating image has vertical anchor PAGE
      And the second floating image has horizontal offset 1828800
      And the second floating image has vertical offset 2743200
      And the third floating image has alt text "Decorative mountain graphic"
      And the third floating image has title "Mountain"


  Scenario: Floating image position dict
    Given a document known to contain three floating images
     Then the document has 3 floating images across its paragraphs
      And the second floating image position dict has the expected keys
