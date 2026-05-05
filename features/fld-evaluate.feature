Feature: Complex-field evaluation (IF / MERGEFIELD / HYPERLINK / formula)
  In order to render data-driven Word documents without launching Word
  As a developer using python-docx
  I need Field.evaluate(context) and Document.evaluate_fields(context)


  Scenario: Field.evaluate substitutes a MERGEFIELD from a context mapping
    Given a simple MERGEFIELD field with name "firstname"
     When I call field.evaluate with {"firstname": "Ada"}
     Then the evaluated result is "Ada"


  Scenario: Field.evaluate on an IF with nested MERGEFIELD picks the true branch
    Given a simple IF field that tests {MERGEFIELD status} equals "active"
     When I call field.evaluate with {"status": "active"}
     Then the evaluated result is "yes"


  Scenario: Field.evaluate on an IF with nested MERGEFIELD picks the false branch
    Given a simple IF field that tests {MERGEFIELD status} equals "active"
     When I call field.evaluate with {"status": "cancelled"}
     Then the evaluated result is "no"


  Scenario: Field.evaluate on a HYPERLINK returns the URL when there is no cached text
    Given a simple HYPERLINK field with url "https://example.com" and no cached text
     When I call field.evaluate with {}
     Then the evaluated result is "https://example.com"


  Scenario: Field.evaluate on a formula field performs the arithmetic
    Given a simple formula field with expression "= 2+3*4"
     When I call field.evaluate with {}
     Then the evaluated result is "14"


  Scenario: Field.evaluate on a PAGE field returns the cached result unchanged
    Given a simple PAGE field with cached text "7"
     When I call field.evaluate with {}
     Then the evaluated result is "7"


  Scenario: Document.evaluate_fields rewrites every field and reports the count
    Given a fresh document containing a MERGEFIELD, an IF, and a formula
     When I call document.evaluate_fields with {"name": "Ada", "status": "active"}
     Then the call returned 3
      And the three fields now read "Ada", "yes", and "5" respectively
