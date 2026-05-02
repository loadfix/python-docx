Feature: Add caption paragraphs to a document
  In order to auto-number figures, tables, and equations using Word's SEQ fields
  As a developer using python-docx
  I need Document.add_caption, Paragraph.add_caption_before, and Paragraph.add_caption_after


  Scenario: Document.add_caption returns a Caption-styled paragraph
    Given a fresh default document
     When I call document.add_caption("A diagram", label="Figure")
     Then the returned paragraph style name is "Caption"
      And the returned paragraph text is "Figure 1: A diagram"


  Scenario Outline: add_caption supports arbitrary labels
    Given a fresh default document
     When I call document.add_caption("<text>", label="<label>")
     Then the returned paragraph text is "<label> 1: <text>"

    Examples: label values
      | label   | text               |
      | Figure  | A diagram          |
      | Table   | A reference table  |
      | Equation| Newton's second    |


  Scenario: Reading back a fixture with two captions
    Given a document having two captions
     Then the document has 2 Caption-styled paragraphs
      And the caption paragraphs contain "Figure 1:" and "Table 1:"


  Scenario: Adding a caption before an existing paragraph
    Given a fresh default document with one paragraph "Intro"
     When I call paragraph.add_caption_before("Preface", label="Figure")
     Then the first paragraph text is "Figure 1: Preface"


  Scenario: Adding a caption after an existing paragraph
    Given a fresh default document with one paragraph "Intro"
     When I call paragraph.add_caption_after("Appendix", label="Figure")
     Then the second paragraph text is "Figure 1: Appendix"
