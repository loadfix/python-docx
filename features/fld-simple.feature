Feature: Simple (w:fldSimple) field codes
  In order to read instruction and result text for one-element field codes
  As a developer using python-docx
  I need a Field proxy over w:fldSimple elements


  Scenario: Paragraph.fields exposes a simple field
    Given a document with a simple DATE field in paragraph 2
     Then paragraph.fields has 1 entry
      And the field is not a complex field


  Scenario: Field.instruction reads the w:instr attribute
    Given a document with a simple DATE field in paragraph 2
     Then field.instruction == "DATE"


  Scenario: Field.type extracts the leading instruction token
    Given a document with a simple DATE field in paragraph 2
     Then field.type == "DATE"


  Scenario: Field.result_text reads cached render text from runs
    Given a document with a simple DATE field in paragraph 2
     Then field.result_text == "2025-01-02"


  Scenario Outline: WD_FIELD_TYPE constants match instruction tokens
    Given the WD_FIELD_TYPE constant <name>
     Then the constant value == "<value>"

    Examples: field-type constants
      | name     | value    |
      | PAGE     | PAGE     |
      | NUMPAGES | NUMPAGES |
      | DATE     | DATE     |
      | TIME     | TIME     |
      | AUTHOR   | AUTHOR   |
      | REF      | REF      |
      | PAGEREF  | PAGEREF  |
      | TOC      | TOC      |


  Scenario: Paragraph.add_simple_field() appends a new simple field
    Given a new empty document
     When I call paragraph.add_simple_field("AUTHOR", "Jane Doe")
     Then paragraph.fields has 1 entry
      And the field is not a complex field
      And field.type == "AUTHOR"
      And field.result_text == "Jane Doe"
