Feature: Read the document's theme part
  In order to inspect the color scheme and font scheme Word applies to the document
  As a developer using python-docx
  I need Document.theme returning a Theme proxy


  Scenario: Document.theme exposes the Office Theme of the default template
    Given a document having the default Office theme
     Then document.theme.name == "Office Theme"


  Scenario Outline: Theme color slots resolve to RGBColor values
    Given a document having the default Office theme
     Then theme.colors.<slot> is a RGBColor

    Examples: theme color slots
      | slot      |
      | dark_1    |
      | light_1   |
      | accent_1  |
      | accent_2  |
      | accent_3  |
      | accent_4  |
      | accent_5  |
      | accent_6  |
      | hyperlink |


  Scenario: Theme colors lookup by OOXML token
    Given a document having the default Office theme
     Then theme.colors["accent1"] is a RGBColor
      And theme.colors["hlink"] is a RGBColor


  Scenario: Theme colors reject unknown tokens
    Given a document having the default Office theme
     Then theme.colors["bogus"] raises KeyError


  Scenario: Theme fonts expose major and minor Latin typefaces
    Given a document having the default Office theme
     Then theme.fonts.major_latin == "Calibri"
      And theme.fonts.minor_latin == "Cambria"
