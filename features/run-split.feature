Feature: Split a run at a character offset
  In order to apply distinct formatting to a sub-range of text inside a run
  As a developer using python-docx
  I need a way to split a run into two runs at a character offset


  Scenario: Run.split() at an interior offset yields two runs with the correct text
    Given a run containing the text "foobar"
     When I split the run at offset 3
     Then the left run text is "foo"
      And the right run text is "bar"
      And the paragraph contains two runs


  Scenario: Run.split() at offset zero yields an empty left run
    Given a run containing the text "foobar"
     When I split the run at offset 0
     Then the left run text is empty
      And the right run text is "foobar"


  Scenario: Run.split() at end offset yields an empty right run
    Given a run containing the text "foobar"
     When I split the run at offset 6
     Then the left run text is "foobar"
      And the right run text is empty


  Scenario: Run.split() preserves run formatting on both sides
    Given a bold italic run containing the text "foobar"
     When I split the run at offset 3
     Then both runs are bold and italic
