Feature: Get and set font color
  In order to work with the color of text
  As a developer using python-docx
  I need a way to get and set the text color


  Scenario Outline: Get font color type
    Given a font having <type> color
     Then font.color.type is <value>

    Examples: Color type settings
      | type    | value |
      | no      | None  |
      | auto    | AUTO  |
      | an RGB  | RGB   |
      | a theme | THEME |


  Scenario Outline: Get font RGB color
    Given a font having <type> color
     Then font.color.rgb is <value>

    Examples: Color type settings
      | type    | value  |
      | no      | None   |
      | auto    | None   |
      | an RGB  | 008000 |
      | a theme | 4f81bd |


  Scenario Outline: Set font RGB color
    Given a font having <type> color
     When I assign <value> to font.color.rgb
     Then font.color.type is <type-value>
     Then font.color.rgb is <rgb-value>

    Examples: Color type settings
      | type    | value  | type-value | rgb-value |
      | no      | f00ba5 | RGB        | f00ba5    |
      | auto    | 2468ac | RGB        | 2468ac    |
      | an RGB  | feeb1e | RGB        | feeb1e    |
      | a theme | 987bac | RGB        | 987bac    |
      | an RGB  | None   | None       | None      |
      | a theme | None   | None       | None      |


  Scenario Outline: Get font theme color
    Given a font having <type> color
     Then font.color.theme_color is <value>

    Examples: Color type settings
      | type    | value    |
      | no      | None     |
      | auto    | None     |
      | an RGB  | None     |
      | a theme | ACCENT_1 |


  Scenario Outline: Set font theme color
    Given a font having <type> color
     When I assign <value> to font.color.theme_color
     Then font.color.type is <type-value>
     Then font.color.theme_color is <theme-value>

    Examples: Color type settings
      | type    | value    | type-value | theme-value |
      | no      | ACCENT_2 | THEME      | ACCENT_2    |
      | auto    | DARK_1   | THEME      | DARK_1      |
      | an RGB  | TEXT_1   | THEME      | TEXT_1      |
      | a theme | LIGHT_2  | THEME      | LIGHT_2     |
      | a theme | None     | None       | None        |
      | an RGB  | None     | None       | None        |


  Scenario: Get font.shading_color when unset
    Given a font
     Then font.shading_color is None


  Scenario Outline: Set and read font.shading_color
    Given a font
     When I assign <value> to font.shading_color
     Then font.shading_color is <expected>

    Examples: run-level background shading colors
      | value   | expected |
      | FF00FF  | FF00FF   |
      | 336699  | 336699   |
      | None    | None     |


  Scenario: Clearing shading_color removes the w:shd element
    Given a font
     When I assign 00FF00 to font.shading_color
      And I assign None to font.shading_color
     Then font.shading_color is None


  Scenario: shading_color is independent of highlight_color
    Given a font
     When I assign BRIGHT_GREEN to font.highlight_color
      And I assign CC0000 to font.shading_color
     Then font.highlight_color is BRIGHT_GREEN
      And font.shading_color is CC0000
