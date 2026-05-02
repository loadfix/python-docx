Feature: Build OMML subscript equations
  In order to author indexed identifiers without hand-writing OMML
  As a developer using python-docx
  I need docx.equations.build_subscript to produce ``m:oMath/m:sSub`` XML


  Scenario: build_subscript with a numeric subscript
    When I call build_subscript("a", "1")
     Then the built equation text is "a1"
      And the built equation raw_xml contains "<m:sSub>"
      And the built equation raw_xml contains "<m:e><m:r><m:t>a</m:t></m:r></m:e>"
      And the built equation raw_xml contains "<m:sub><m:r><m:t>1</m:t></m:r></m:sub>"
      And the built equation is_display_mode is False


  Scenario: Round-trip a subscript through Paragraph.add_equation
    Given a fresh default document
     When I add a subscript equation "a" "n" to a new paragraph
     Then the paragraph has 1 equation
      And the paragraph first equation text is "an"
      And the paragraph first equation raw_xml contains "<m:sSub>"


  Scenario: Chained subscripts in a single paragraph
    Given a fresh default document
     When I add a subscript equation "a" "i" to a new paragraph
      And I add a subscript equation "b" "j" to the same paragraph
     Then the paragraph has 2 equations
      And the paragraph first equation text is "ai"
      And the paragraph second equation text is "bj"
      And every paragraph equation raw_xml contains "<m:sSub>"


  Scenario: Reading back chained subscripts from a fixture
    Given a document having chained subscript equations
     Then the document has 3 subscript equations
      And the first subscript equation text is "a1"
      And the second subscript equation text is "ai"
      And the third subscript equation text is "bj"
