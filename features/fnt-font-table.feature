Feature: Enumerate fonts referenced by a document
  In order to inspect which fonts Word embedded metadata for
  As a developer using python-docx
  I need Document.font_table returning a FontTable collection


  Scenario: Document.font_table exposes the default template's font entries
    Given a document having a font table
     Then document.font_table is not None
      And len(document.font_table) is at least 4
      And "Calibri" is in document.font_table


  Scenario: FontTable lookup by name returns a FontMetadata
    Given a document having a font table
     Then document.font_table["Calibri"].name == "Calibri"
      And document.font_table["Calibri"].panose has length 20


  Scenario: FontTable.get() returns None for missing fonts
    Given a document having a font table
     Then document.font_table.get("NoSuchFont") is None


  Scenario: FontTable lookup with unknown name raises KeyError
    Given a document having a font table
     Then document.font_table["NoSuchFont"] raises KeyError


  Scenario: Iterating a FontTable yields FontMetadata instances
    Given a document having a font table
     Then iterating document.font_table yields only FontMetadata objects
