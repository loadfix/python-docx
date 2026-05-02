Feature: Build OMML superscript equations
  In order to author common math idioms without hand-writing OMML
  As a developer using python-docx
  I need docx.equations.build_superscript to produce ``m:oMath/m:sSup`` XML


  Scenario: build_superscript with a numeric exponent
    When I call build_superscript("x", "2")
     Then the built equation text is "x2"
      And the built equation raw_xml contains "<m:sSup>"
      And the built equation raw_xml contains "<m:e><m:r><m:t>x</m:t></m:r></m:e>"
      And the built equation raw_xml contains "<m:sup><m:r><m:t>2</m:t></m:r></m:sup>"
      And the built equation is_display_mode is False


  Scenario: build_superscript with an identifier exponent
    When I call build_superscript("e", "x")
     Then the built equation text is "ex"
      And the built equation raw_xml contains "<m:sSup>"
      And the built equation raw_xml contains "<m:e><m:r><m:t>e</m:t></m:r></m:e>"
      And the built equation raw_xml contains "<m:sup><m:r><m:t>x</m:t></m:r></m:sup>"


  Scenario: Round-trip a superscript through Paragraph.add_equation
    Given a fresh default document
     When I add a superscript equation "x" "2" to a new paragraph
     Then the paragraph has 1 equation
      And the paragraph first equation text is "x2"
      And the paragraph first equation raw_xml contains "<m:sSup>"


  Scenario: Reading back both superscripts from a fixture
    Given a document having two superscript equations
     Then the document has 2 superscript equations
      And the first superscript equation text is "x2"
      And the second superscript equation text is "ex"
      And every superscript equation raw_xml contains "<m:sSup>"
