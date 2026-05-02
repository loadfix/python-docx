Feature: Build OMML radical expressions with docx.equations.build_radical
  In order to emit OMML radicals (square roots and nth roots) without hand-
  writing XML
  As a developer using python-docx
  I need docx.equations.build_radical(expr_text[, degree_text])


  Scenario: build_radical with no degree produces a bare square root
    When I build a radical with expr "x" and no degree
    Then the OMML contains an "m:rad" element
     And the OMML contains an empty "m:deg" element
     And the equation text is "x"


  Scenario: build_radical with a degree produces an nth root
    When I build a radical with expr "x" and degree "3"
    Then the OMML contains an "m:rad" element
     And the OMML contains a populated "m:deg" element
     And the equation text is "3x"


  Scenario Outline: build_radical handles assorted radicand and degree values
    When I build a radical with expr "<expr>" and degree "<degree>"
    Then the equation text is "<text>"

    Examples: nth-root inputs
      | expr   | degree | text     |
      | a      | 2      | 2a       |
      | n + 1  | 4      | 4n + 1   |
      | 16     | 3      | 316      |


  Scenario: Nested radical round-trips through a paragraph
    Given a document having a radical-equation fixture
     Then the document has 3 radical equations
      And the third radical has a nested "m:rad" descendant


  Scenario: Appending a radical equation to a paragraph
    Given a fresh default document
     When I append a radical equation with expr "z" and degree "5" to a new paragraph
     Then the paragraph has 1 equation
      And the appended equation text is "5z"
      And the appended equation is not display mode
