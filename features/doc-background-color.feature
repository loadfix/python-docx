Feature: Document.background_color
  In order to control the page background colour of a document
  As a developer using python-docx
  I need a read/write Document.background_color property


  Scenario: Get background_color when not set
    Given a blank document
     Then document.background_color is None


  Scenario: Get background_color from a document with a colour set
    Given a document having a background color
     Then document.background_color is RGBColor(FF, A5, 00)


  Scenario Outline: Set background_color from scratch
    Given a blank document
     When I assign RGBColor(<hex>) to document.background_color
     Then document.background_color is RGBColor(<hex>)

    Examples: background colour hex values
      | hex        |
      | 00, 80, FF |
      | 7F, 7F, 7F |


  Scenario: Overwrite an existing background_color
    Given a document having a background color
     When I assign RGBColor(00, 80, 00) to document.background_color
     Then document.background_color is RGBColor(00, 80, 00)


  Scenario: Clear background_color by assigning None
    Given a document having a background color
     When I assign None to document.background_color
     Then document.background_color is None


  Scenario: background_color round-trips through save/load
    Given a blank document
     When I assign RGBColor(AB, CD, EF) to document.background_color
      And I save and reload the document to scratch
     Then document.background_color is RGBColor(AB, CD, EF)
