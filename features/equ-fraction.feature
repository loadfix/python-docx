Feature: Build and read fraction equations
  In order to render stacked fractions with a horizontal bar in OMML
  As a developer using python-docx
  I need the build_fraction helper to wrap numerator and denominator
  identifiers in an m:f element and I need to read the result back


  Scenario: Document.equations yields a list of Equation objects
    Given a document having fraction equations
     Then document.equations is a list of 2 Equation objects


  Scenario: Reading the flattened text of a simple a/b fraction
    Given a document having fraction equations
     When I assign equation = the first equation in the document
     Then equation.text == "ab"


  Scenario: Reading the flattened text of a compound (x+1)/y fraction
    Given a document having fraction equations
     When I assign equation = the second equation in the document
     Then equation.text == "x+1y"


  Scenario: A fraction equation is inline, not display-mode
    Given a document having fraction equations
     When I assign equation = the first equation in the document
     Then equation.is_display_mode is False


  Scenario: The raw OMML of a fraction equation includes m:f
    Given a document having fraction equations
     When I assign equation = the first equation in the document
     Then equation.raw_xml contains "<m:f>"
      And equation.raw_xml contains "<m:num>"
      And equation.raw_xml contains "<m:den>"


  Scenario: build_fraction returns a parseable m:oMath fragment with m:f
    Given a fresh document
     When I assign xml = build_fraction("p", "q")
     Then xml starts with "<m:oMath"
      And xml ends with "</m:oMath>"
      And xml contains "<m:f>"
      And xml contains "<m:num><m:r><m:t>p</m:t></m:r></m:num>"
      And xml contains "<m:den><m:r><m:t>q</m:t></m:r></m:den>"


  Scenario: build_fraction emits a horizontal bar type
    Given a fresh document
     When I assign xml = build_fraction("a", "b")
     Then xml contains a bar-type fraction property


  Scenario: Paragraph.add_equation with build_fraction appends an Equation
    Given a fresh document
     When I append a fraction equation with numerator "m" and denominator "n" to a new paragraph
     Then the paragraph has 1 equation
      And document.equations is a list of 1 Equation objects
      And the appended equation.text == "mn"
