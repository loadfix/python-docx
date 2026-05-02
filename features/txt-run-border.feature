Feature: Read/write run-level border properties
  In order to draw a border around a run of text
  As a python-docx developer
  I need the Font.border_* properties to expose w:rPr/w:bdr


  Scenario Outline: Get border style
    Given a run from txt-run-border paragraph <idx>
     Then font.border_style is <value>

    Examples: border-style read values
      | idx | value                  |
      | 0   | None                   |
      | 1   | WD_BORDER_STYLE.SINGLE |
      | 2   | WD_BORDER_STYLE.DASHED |


  Scenario Outline: Get border color
    Given a run from txt-run-border paragraph <idx>
     Then font.border_color is <value>

    Examples: border-color read values (auto normalised to None)
      | idx | value  |
      | 0   | None   |
      | 1   | FF0000 |
      | 2   | None   |


  Scenario Outline: Get border width and space
    Given a run from txt-run-border paragraph <idx>
     Then font.border_width is <width>
      And font.border_space is <space>

    Examples: border width/space read values
      | idx | width  | space |
      | 0   | None   | None  |
      | 1   | Pt(1.5)| Pt(4) |


  Scenario: Set a border from scratch
    Given a run from txt-run-border paragraph 0
     When I set font.border_style to WD_BORDER_STYLE.DOUBLE
      And I set font.border_color to 00FF00
      And I set font.border_width to Pt(1)
      And I set font.border_space to Pt(2)
     Then font.border_style is WD_BORDER_STYLE.DOUBLE
      And font.border_color is 00FF00
      And font.border_width is Pt(1)
      And font.border_space is Pt(2)


  Scenario: Clear individual border attributes
    Given a run from txt-run-border paragraph 1
     When I set font.border_color to None
      And I set font.border_space to None
     Then font.border_color is None
      And font.border_space is None
      And font.border_style is WD_BORDER_STYLE.SINGLE


  Scenario: Remove the whole border
    Given a run from txt-run-border paragraph 1
     When I call font.remove_border()
     Then font.border_style is None
      And font.border_color is None
      And font.border_width is None
      And font.border_space is None
