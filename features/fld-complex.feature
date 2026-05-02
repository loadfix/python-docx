Feature: Complex (w:fldChar) field codes
  In order to read instruction and result text for multi-run field codes
  As a developer using python-docx
  I need a Field proxy over the w:fldChar begin/separate/end sequence


  Scenario: Paragraph.fields exposes a complex field
    Given a document with a complex PAGE field in paragraph 3
     Then paragraph.fields has 1 entry
      And the field is a complex field


  Scenario: Field.instruction concatenates w:instrText between begin and separate
    Given a document with a complex PAGE field in paragraph 3
     Then field.instruction == "PAGE"


  Scenario: Field.type extracts the leading instruction token
    Given a document with a complex PAGE field in paragraph 3
     Then field.type == "PAGE"


  Scenario: Field.result_text concatenates runs between separate and end
    Given a document with a complex PAGE field in paragraph 3
     Then field.result_text == "7"


  Scenario: Paragraph.add_complex_field() appends a new complex field
    Given a new empty document
     When I call paragraph.add_complex_field("NUMPAGES", "12")
     Then paragraph.fields has 1 entry
      And the field is a complex field
      And field.type == "NUMPAGES"
      And field.result_text == "12"


  Scenario: Field.update_result_text() rewrites the rendered result region
    Given a document with a complex PAGE field in paragraph 3
     When I call field.update_result_text("42")
     Then field.result_text == "42"


  Scenario: Iterating fields across all paragraphs
    Given the fld-has-fields document
     Then iterating every paragraph's fields yields 3 Field objects
