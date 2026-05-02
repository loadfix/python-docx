Feature: Add content to a run
  In order to populate a run with varied content
  As a developer using python-docx
  I need a way to add each of the run content elements to a run

  Scenario: Add a tab
    Given a run
     When I add a tab
     Then the tab appears at the end of the run

  Scenario: Assign mixed text to text property
    Given a run
     When I assign mixed text to the text property
     Then run.text contains the text content of the run


  Scenario: Add a symbol from a named font
    Given a run
     When I add symbol 0xF0E0 from font Wingdings
     Then the last item in the run is a w:sym element
      And the run contains 1 symbol
      And the first symbol has char_code 0xF0E0
      And the first symbol has char_hex F0E0
      And the first symbol has font Wingdings


  Scenario Outline: add_symbol accepts int and hex-string char codes
    Given a run
     When I add symbol <char_code> from font Symbol
     Then the first symbol has char_hex <expected>

    Examples: char_code input forms
      | char_code | expected |
      | 0xF0E0    | F0E0     |
      | 240       | 00F0     |
      | F0E1      | F0E1     |
      | 0xf0e2    | F0E2     |


  Scenario: Enumerate multiple symbols
    Given a run
     When I add symbol 0xF0E0 from font Wingdings
      And I add symbol 0xF0E1 from font Wingdings
      And I add symbol 0xF0E2 from font Wingdings2
     Then the run contains 3 symbols


  Scenario: Delete a symbol
    Given a run
     When I add symbol 0xF0E0 from font Wingdings
      And I add symbol 0xF0E1 from font Wingdings
      And I delete the first symbol
     Then the run contains 1 symbol
      And the first symbol has char_hex F0E1
