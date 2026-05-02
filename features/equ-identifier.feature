Feature: Build and read identifier equations
  In order to attach a single-letter mathematical identifier to running prose
  As a developer using python-docx
  I need the build_identifier helper to produce a well-formed m:oMath element
  and a read path that exposes its text and raw OMML


  Scenario: Document.equations yields a list of Equation objects
    Given a document having identifier equations
     Then document.equations is a list of 2 Equation objects


  Scenario: Reading the text of an identifier equation
    Given a document having identifier equations
     When I assign equation = the first equation in the document
     Then equation.text == "x"


  Scenario: Reading the text of a Greek chi identifier equation
    Given a document having identifier equations
     When I assign equation = the second equation in the document
     Then equation.text == "χ"


  Scenario: An identifier equation is inline, not display-mode
    Given a document having identifier equations
     When I assign equation = the first equation in the document
     Then equation.is_display_mode is False


  Scenario: The raw OMML of an identifier equation includes m:r and m:t
    Given a document having identifier equations
     When I assign equation = the first equation in the document
     Then equation.raw_xml contains "<m:r>"
      And equation.raw_xml contains "<m:t>"


  Scenario: build_identifier returns a parseable m:oMath fragment
    Given a fresh document
     When I assign xml = build_identifier("y")
     Then xml starts with "<m:oMath"
      And xml ends with "</m:oMath>"
      And xml contains "<m:t>y</m:t>"


  Scenario: build_identifier XML-escapes its text argument
    Given a fresh document
     When I assign xml = build_identifier("a<b")
     Then xml contains "<m:t>a&lt;b</m:t>"


  Scenario: Paragraph.add_equation with build_identifier appends an Equation
    Given a fresh document
     When I append an identifier equation for "z" to a new paragraph
     Then the paragraph has 1 equation
      And document.equations is a list of 1 Equation objects
      And the appended equation.text == "z"
