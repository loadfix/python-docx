Feature: Read ink annotations from a document
  In order to inspect stylus-authored ink annotations
  As a developer using python-docx
  I need read access to ``w:contentPart`` references and their InkML parts


  Scenario: Enumerate ink annotations at the document level
    Given a document known to contain two ink annotations
     Then the document exposes 2 ink annotations
      And the ink annotations have partnames "/word/ink/ink1.xml" and "/word/ink/ink2.xml"


  Scenario Outline: Stroke count per annotation
    Given a document known to contain two ink annotations
     Then the ink annotation at partname "<partname>" has <count> strokes

    Examples: stroke counts
      | partname              | count |
      | /word/ink/ink1.xml    |   2   |
      | /word/ink/ink2.xml    |   1   |


  Scenario: Ink annotations accessible via containing paragraph
    Given a document known to contain two ink annotations
     Then the first paragraph carrying ink has 1 annotation
      And the second paragraph carrying ink has 1 annotation
      And each annotation's paragraph is the paragraph that contains it


  Scenario: Document without ink has an empty annotations list
    Given a pristine empty document
     Then the document exposes 0 ink annotations
