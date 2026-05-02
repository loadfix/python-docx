Feature: Build and apply custom numbering definitions
  In order to attach ordered and bulleted lists to paragraphs with custom formatting
  As a developer using python-docx
  I need Numbering.add_numbering_definition and NumberingDefinition.apply_to


  Scenario: Document.numbering exposes definitions from the numbering part
    Given a document having a custom numbering definition
     Then document.numbering has at least 1 definition
      And the last numbering definition has 3 levels


  Scenario Outline: Level properties reflect the source w:lvl element
    Given a document having a custom numbering definition
     Then level <ilvl> of the last definition has number_format == <fmt>
      And level <ilvl> of the last definition has text == "<text>"
      And level <ilvl> of the last definition has start == <start>

    Examples: per-level metadata
      | ilvl | fmt         | text | start |
      | 0    | DECIMAL     | %1.  | 5     |
      | 1    | LOWER_LETTER| %2)  | 1     |


  Scenario: Paragraphs with an attached numbering have a w:numPr element
    Given a document having a custom numbering definition
     Then the first three paragraphs have a w:numPr child


  Scenario: Adding a single-level bullet definition
    Given a fresh default document
     When I add a single-level bullet numbering definition
     Then document.numbering has at least 1 definition
      And the last numbering definition has 1 level


  Scenario: Applying a numbering definition to a paragraph
    Given a fresh default document with one paragraph
     When I add a single-level decimal numbering definition
      And I apply the last numbering definition to the paragraph at level 0
     Then the paragraph has a w:numPr child


  Scenario: apply_to raises for out-of-range levels
    Given a fresh default document with one paragraph
     When I add a single-level decimal numbering definition
     Then applying the definition to the paragraph at level 9 raises ValueError
