Feature: Read/write East Asian layout and typography
  In order to format Japanese, Chinese, or Korean text
  As a python-docx developer
  I need access to w:eastAsianLayout and the kinsoku/word-wrap paragraph toggles


  Scenario: font has no east-asian layout when w:eastAsianLayout is absent
    Given a run from txt-east-asian paragraph 0
     Then the run has no east-asian layout


  Scenario Outline: font exposes an EastAsianLayout when w:eastAsianLayout is present
    Given a run from txt-east-asian paragraph <idx>
     Then the run has an east-asian layout

    Examples: east_asian_layout presence
      | idx |
      | 1   |
      | 2   |


  Scenario Outline: Get east-asian layout attributes
    Given a run from txt-east-asian paragraph <idx>
     Then font.east_asian_layout.two_lines_in_one is <combine>
      And font.east_asian_layout.vertical_alignment is <vert>
      And font.east_asian_layout.compressed is <compressed>

    Examples: east_asian_layout attribute values
      | idx | combine | vert | compressed |
      | 1   | True    | None | None       |
      | 2   | None    | True | True       |


  Scenario: Update an east-asian layout in place
    Given a run from txt-east-asian paragraph 1
     When I call font.set_east_asian_layout compressed=True
     Then font.east_asian_layout.compressed is True
      And font.east_asian_layout.two_lines_in_one is True


  Scenario: Create an east-asian layout on a bare run
    Given a run from txt-east-asian paragraph 0
     When I call font.set_east_asian_layout two_lines_in_one=True id=5
     Then font.east_asian_layout.two_lines_in_one is True
      And font.east_asian_layout.id is 5


  Scenario: Remove the east-asian layout element
    Given a run from txt-east-asian paragraph 1
     When I call font.remove_east_asian_layout()
     Then the run has no east-asian layout


  Scenario Outline: Read paragraph kinsoku value
    Given a paragraph format from txt-east-asian paragraph <idx>
     Then paragraph_format.kinsoku is <value>

    Examples: kinsoku read values
      | idx | value |
      | 0   | None  |
      | 1   | False |
      | 2   | True  |


  Scenario Outline: Read paragraph word_wrap value
    Given a paragraph format from txt-east-asian paragraph <idx>
     Then paragraph_format.word_wrap is <value>

    Examples: word_wrap read values
      | idx | value |
      | 0   | None  |
      | 1   | None  |
      | 2   | False |


  Scenario Outline: Set paragraph kinsoku and word_wrap
    Given a paragraph format from txt-east-asian paragraph <idx>
     When I assign <value> to paragraph_format.<prop_name>
     Then paragraph_format.<prop_name> is <value>

    Examples: kinsoku/word_wrap assignments
      | idx | prop_name | value |
      | 0   | kinsoku   | True  |
      | 0   | kinsoku   | False |
      | 1   | kinsoku   | None  |
      | 0   | word_wrap | False |
      | 2   | word_wrap | True  |
      | 2   | word_wrap | None  |
