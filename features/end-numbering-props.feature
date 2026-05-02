Feature: Document-level endnote numbering and position properties
  In order to configure how endnotes are numbered and positioned in a document
  As a developer using python-docx
  I need access to the EndnoteProperties proxy over the w:endnotePr element


  Scenario: Document.endnote_properties is None when no w:endnotePr exists
    Given a document having no endnotes part
     Then document.endnote_properties is None


  Scenario: Document.add_endnote_properties() returns an EndnoteProperties object
    Given a document having no endnotes part
     When I assign props = document.add_endnote_properties()
     Then props is an EndnoteProperties object
      And document.endnote_properties is an EndnoteProperties object


  Scenario: EndnoteProperties.number_format reads the document setting
    Given a document having an endnotes part
     Then document.endnote_properties.number_format == WD_NUMBER_FORMAT.LOWER_ROMAN


  Scenario Outline: EndnoteProperties.number_format can be updated
    Given a document having an endnotes part
     When I assign props.number_format = WD_NUMBER_FORMAT.<value>
     Then document.endnote_properties.number_format == WD_NUMBER_FORMAT.<value>

    Examples: supported number formats
      | value        |
      | ARABIC       |
      | UPPER_ROMAN  |
      | LOWER_ROMAN  |
      | UPPER_LETTER |
      | LOWER_LETTER |
      | CHICAGO      |


  Scenario: EndnoteProperties.number_format can be cleared to None
    Given a document having an endnotes part
     When I assign props.number_format = None
     Then document.endnote_properties.number_format is None


  Scenario: EndnoteProperties.restart_rule reads the document setting
    Given a document having an endnotes part
     Then document.endnote_properties.restart_rule == WD_FOOTNOTE_RESTART.CONTINUOUS


  Scenario Outline: EndnoteProperties.restart_rule can be updated
    Given a document having an endnotes part
     When I assign props.restart_rule = WD_FOOTNOTE_RESTART.<value>
     Then document.endnote_properties.restart_rule == WD_FOOTNOTE_RESTART.<value>

    Examples: supported restart rules
      | value        |
      | CONTINUOUS   |
      | EACH_SECTION |


  Scenario: EndnoteProperties.position reads the document setting
    Given a document having an endnotes part
     Then document.endnote_properties.position == WD_ENDNOTE_POSITION.END_OF_DOCUMENT


  Scenario Outline: EndnoteProperties.position can be updated
    Given a document having an endnotes part
     When I assign props.position = WD_ENDNOTE_POSITION.<value>
     Then document.endnote_properties.position == WD_ENDNOTE_POSITION.<value>

    Examples: supported positions
      | value           |
      | END_OF_DOCUMENT |
      | END_OF_SECTION  |


  Scenario: EndnoteProperties.start_number reads the document setting
    Given a document having an endnotes part
     Then document.endnote_properties.start_number == 1


  Scenario: EndnoteProperties.start_number can be updated
    Given a document having an endnotes part
     When I assign props.start_number = 7
     Then document.endnote_properties.start_number == 7
