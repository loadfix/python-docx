Feature: Cross-reference (REF / PAGEREF) field resolution
  In order to replace REF fields with the current text of the referenced bookmark
  As a developer using python-docx
  I need Field.resolve() and Document.resolve_cross_references()


  Scenario: Field.resolve() returns bookmark text for a REF field
    Given a document with a REF field pointing at the bookmark "FavouriteValue"
     Then field.resolve(document) == "The quoted value is forty-two."


  Scenario: Field.resolve() returns cached text for a PAGE field
    Given a document with a complex PAGE field in paragraph 3
     Then field.resolve(document) == "7"


  Scenario: PAGEREF fields fall back to "?" when no cached result is present
    Given a new empty document with an unresolved PAGEREF field
     Then field.resolve(document) == "?"


  Scenario: Document.resolve_cross_references() rewrites REF result text
    Given the fld-has-fields document
     When I call document.resolve_cross_references()
     Then the call returned 0
      And the REF field in paragraph 5 still reads "The quoted value is forty-two."


  Scenario: Document.resolve_cross_references() reports updates when REF content drifts
    Given the fld-has-fields document with a stale REF result
     When I call document.resolve_cross_references()
     Then the call returned 1
      And the REF field in paragraph 5 still reads "The quoted value is forty-two."
