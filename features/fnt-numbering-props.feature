Feature: Document-level footnote numbering properties
  In order to control how footnotes are numbered and positioned
  As a developer using python-docx
  I need access to FootnoteProperties on the document


  Scenario Outline: Document.footnote_properties access
    Given a document <with-or-without> footnote properties
     Then document.footnote_properties is <result>

    Examples: footnote_properties access cases
      | with-or-without | result                      |
      | with            | a FootnoteProperties object |
      | without         | None                        |


  Scenario: FootnoteProperties.number_format getter
    Given a document with footnote properties
     Then document.footnote_properties.number_format == WD_NUMBER_FORMAT.LOWER_ROMAN


  Scenario: FootnoteProperties.start_number getter
    Given a document with footnote properties
     Then document.footnote_properties.start_number == 7


  Scenario: FootnoteProperties.restart_rule getter
    Given a document with footnote properties
     Then document.footnote_properties.restart_rule == WD_FOOTNOTE_RESTART.EACH_SECTION


  Scenario: FootnoteProperties.position getter
    Given a document with footnote properties
     Then document.footnote_properties.position == WD_FOOTNOTE_POSITION.BENEATH_TEXT


  Scenario: Document.add_footnote_properties() on a document without them
    Given a document without footnote properties
     When I call document.add_footnote_properties()
     Then document.footnote_properties is a FootnoteProperties object
      And document.footnote_properties.number_format is None
      And document.footnote_properties.start_number is None
      And document.footnote_properties.restart_rule is None
      And document.footnote_properties.position is None


  Scenario: Set FootnoteProperties.number_format
    Given a document without footnote properties
     When I call document.add_footnote_properties()
      And I assign WD_NUMBER_FORMAT.UPPER_ROMAN to footnote_properties.number_format
     Then document.footnote_properties.number_format == WD_NUMBER_FORMAT.UPPER_ROMAN


  Scenario: Set FootnoteProperties.start_number
    Given a document without footnote properties
     When I call document.add_footnote_properties()
      And I assign 5 to footnote_properties.start_number
     Then document.footnote_properties.start_number == 5


  Scenario: Set FootnoteProperties.restart_rule
    Given a document without footnote properties
     When I call document.add_footnote_properties()
      And I assign WD_FOOTNOTE_RESTART.EACH_PAGE to footnote_properties.restart_rule
     Then document.footnote_properties.restart_rule == WD_FOOTNOTE_RESTART.EACH_PAGE


  Scenario: Set FootnoteProperties.position
    Given a document without footnote properties
     When I call document.add_footnote_properties()
      And I assign WD_FOOTNOTE_POSITION.BOTTOM_OF_PAGE to footnote_properties.position
     Then document.footnote_properties.position == WD_FOOTNOTE_POSITION.BOTTOM_OF_PAGE
