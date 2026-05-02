Feature: Embed a TrueType font into a document
  In order to ship a document with a font that the reader's system may not have
  As a developer using python-docx
  I need FontTable.add_embedded_font() plus round-trip preservation


  Scenario: add_embedded_font registers a new font entry
    Given a document with no font table
     When I call document.font_table_or_new.add_embedded_font on a sample font
     Then document.font_table is not None
      And the font table has one embedded-regular entry


  Scenario: embedded font binaries round-trip through save/load
    Given a document with no font table
     When I call document.font_table_or_new.add_embedded_font on a sample font
      And I save and reopen the font-embed document
     Then document.font_table is not None
      And the font table has one embedded-regular entry
      And the embedded font binary matches the original
