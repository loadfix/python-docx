Feature: Get or set font properties
  In order to customize the character formatting of text in a document
  As a python-docx developer
  I need a set of read/write properties on the Font object


  Scenario Outline: Get highlight color
    Given a font having <color> highlighting 
     Then font.highlight_color is <value>

    Examples: font.highlight_color values
      | color           | value        |
      | no              | None         |
      | yellow          | YELLOW       |
      | bright green    | BRIGHT_GREEN |


  Scenario Outline: Set highlight color
    Given a font having <color> highlighting
     When I assign <value> to font.highlight_color
     Then font.highlight_color is <value>

    Examples: font.highlight_color values
      | color           | value        |
      | no              | YELLOW       |
      | yellow          | None         |
      | bright green    | BRIGHT_GREEN |
      
      
  Scenario Outline: Get typeface name
    Given a font having typeface name <name>
     Then font.name is <value>

    Examples: font.name values
      | name          | value        |
      | not specified | None         |
      | Avenir Black  | Avenir Black |


  Scenario Outline: Set typeface name
    Given a font having typeface name <name>
     When I assign <value> to font.name
     Then font.name is <value>

    Examples: font.name values
      | name          | value        |
      | not specified | Avenir Black |
      | Avenir Black  | Calibri      |
      | Avenir Black  | None         |


  Scenario Outline: Get font size
    Given a font of size <size>
     Then font.size is <value>

    Examples: font.size values
      | size        | value  |
      | unspecified | None   |
      | 14 pt       | 177800 |


  Scenario Outline: Set font size
    Given a font of size <size>
     When I assign <value> to font.size
     Then font.size is <value>

    Examples: font.size post-assignment values
      | size        | value  |
      | unspecified | 177800 |
      | 14 pt       | 228600 |
      | 18 pt       | None   |


  Scenario: Get font color object
    Given a font
     Then font.color is a ColorFormat object


  Scenario Outline: Get font underline value
    Given a font having <underline-type> underline
     Then font.underline is <value>

    Examples: font underline values
      | underline-type | value               |
      | inherited      | None                |
      | no             | False               |
      | single         | True                |
      | double         | WD_UNDERLINE.DOUBLE |


  Scenario Outline: Change font underline
    Given a font having <underline-type> underline
     When I assign <new-value> to font.underline
     Then font.underline is <expected-value>

    Examples: underline property values
      | underline-type | new-value           | expected-value      |
      | inherited      | True                | True                |
      | inherited      | False               | False               |
      | inherited      | None                | None                |
      | inherited      | WD_UNDERLINE.SINGLE | True                |
      | inherited      | WD_UNDERLINE.DOUBLE | WD_UNDERLINE.DOUBLE |
      | single         | None                | None                |
      | single         | True                | True                |
      | single         | False               | False               |
      | single         | WD_UNDERLINE.SINGLE | True                |
      | single         | WD_UNDERLINE.DOUBLE | WD_UNDERLINE.DOUBLE |


  Scenario Outline: Get font sub/superscript value
    Given a font having <vertAlign-state> vertical alignment
     Then font.subscript is <sub-value>
      And font.superscript is <super-value>

    Examples: font sub/superscript values
      | vertAlign-state | sub-value | super-value |
      | inherited       | None      | None        |
      | subscript       | True      | False       |
      | superscript     | False     | True        |


  Scenario Outline: Change font sub/superscript
    Given a font having <vertAlign-state> vertical alignment
     When I assign <value> to font.<name>script
     Then font.<name-2>script is <expected-value>

    Examples: value of sub/superscript after assignment
      | vertAlign-state | name  | value | name-2  | expected-value |
      | inherited       | sub   | True  |  sub    | True           |
      | inherited       | sub   | True  |  super  | False          |
      | inherited       | sub   | False |  sub    | None           |
      | inherited       | super | True  |  super  | True           |
      | inherited       | super | True  |  sub    | False          |
      | inherited       | super | False |  super  | None           |
      | subscript       | sub   | True  |  sub    | True           |
      | subscript       | sub   | False |  sub    | None           |
      | subscript       | sub   | None  |  sub    | None           |
      | subscript       | super | True  |  sub    | False          |
      | subscript       | super | False |  sub    | True           |
      | subscript       | super | None  |  sub    | None           |
      | superscript     | super | True  |  super  | True           |
      | superscript     | super | False |  super  | None           |
      | superscript     | super | None  |  super  | None           |
      | superscript     | sub   | True  |  super  | False          |
      | superscript     | sub   | False |  super  | True           |
      | superscript     | sub   | None  |  super  | None           |


  Scenario Outline: Apply boolean property to a run
    Given a run
     When I assign True to its <boolean_prop_name> property
     Then the run appears in <boolean_prop_name> unconditionally

    Examples: Boolean run properties
      | boolean_prop_name |
      | all_caps          |
      | bold              |
      | complex_script    |
      | cs_bold           |
      | cs_italic         |
      | double_strike     |
      | emboss            |
      | hidden            |
      | italic            |
      | imprint           |
      | math              |
      | no_proof          |
      | outline           |
      | rtl               |
      | shadow            |
      | small_caps        |
      | snap_to_grid      |
      | spec_vanish       |
      | strike            |
      | web_hidden        |


  Scenario Outline: Set <boolean_prop_name> off unconditionally
    Given a run
     When I assign False to its <boolean_prop_name> property
     Then the run appears without <boolean_prop_name> unconditionally

    Examples: Boolean run properties
      | boolean_prop_name |
      | all_caps          |
      | bold              |
      | complex_script    |
      | cs_bold           |
      | cs_italic         |
      | double_strike     |
      | emboss            |
      | hidden            |
      | italic            |
      | imprint           |
      | math              |
      | no_proof          |
      | outline           |
      | rtl               |
      | shadow            |
      | small_caps        |
      | snap_to_grid      |
      | spec_vanish       |
      | strike            |
      | web_hidden        |


  Scenario Outline: Remove boolean property from a run
    Given a run having <boolean_prop_name> set on
     When I assign None to its <boolean_prop_name> property
     Then the run appears with its inherited <boolean_prop_name> setting

    Examples: Boolean run properties
      | boolean_prop_name |
      | all_caps          |
      | bold              |
      | complex_script    |
      | cs_bold           |
      | cs_italic         |
      | double_strike     |
      | emboss            |
      | hidden            |
      | italic            |
      | imprint           |
      | math              |
      | no_proof          |
      | outline           |
      | rtl               |
      | shadow            |
      | small_caps        |
      | snap_to_grid      |
      | spec_vanish       |
      | strike            |
      | web_hidden        |


  Scenario: Get font.kerning when unset
    Given a font
     Then font.kerning is None


  Scenario Outline: Set and get font.kerning
    Given a font
     When I assign <value> to font.kerning
     Then font.kerning is <expected>

    Examples: font.kerning values (half-points via Pt)
      | value   | expected |
      | Pt(10)  | Pt(10)   |
      | Pt(8)   | Pt(8)    |
      | None    | None     |


  Scenario: Get font.character_spacing when unset
    Given a font
     Then font.character_spacing is None


  Scenario Outline: Set and get font.character_spacing
    Given a font
     When I assign <value> to font.character_spacing
     Then font.character_spacing is <expected>

    Examples: font.character_spacing values (twentieths of a point via Pt)
      | value    | expected |
      | Pt(1)    | Pt(1)    |
      | Pt(-0.5) | Pt(-0.5) |
      | None     | None     |


  Scenario: Get font.right_to_left default
    Given a font
     Then font.right_to_left is False


  Scenario Outline: Set font.right_to_left
    Given a font
     When I assign <value> to font.right_to_left
     Then font.right_to_left is <expected>

    Examples: right-to-left toggle
      | value | expected |
      | True  | True     |
      | False | False    |
      | None  | False    |


  Scenario: Get font.language when unset
    Given a font
     Then font.language is None
      And font.east_asian_language is None
      And font.bidi_language is None


  Scenario Outline: Set language tags
    Given a font
     When I assign <value> to font.<prop_name>
     Then font.<prop_name> is <expected>

    Examples: language-tag assignments
      | prop_name           | value | expected |
      | language            | en-US | en-US    |
      | language            | fr-FR | fr-FR    |
      | language            | None  | None     |
      | east_asian_language | ja-JP | ja-JP    |
      | east_asian_language | None  | None     |
      | bidi_language       | ar-SA | ar-SA    |
      | bidi_language       | None  | None     |


  Scenario: Remove the entire w:lang element
    Given a font
     When I assign en-US to font.language
      And I assign ja-JP to font.east_asian_language
      And I assign ar-SA to font.bidi_language
      And I call font.remove_language()
     Then font.language is None
      And font.east_asian_language is None
      And font.bidi_language is None


  Scenario: Get font.name_far_east when unset
    Given a font
     Then font.name_far_east is None
      And font.name_east_asia is None


  Scenario Outline: Set font.name_far_east
    Given a font
     When I assign <value> to font.<prop_name>
     Then font.name_far_east is <expected>
      And font.name_east_asia is <expected>

    Examples: East-Asian typeface assignments (both spellings are aliased)
      | prop_name      | value      | expected   |
      | name_far_east  | MS Mincho  | MS Mincho  |
      | name_east_asia | SimSun     | SimSun     |
      | name_far_east  | None       | None       |


  Scenario: name_east_asia is a writable alias for name_far_east
    Given a font
     When I assign MS Gothic to font.name_east_asia
     Then font.name_far_east is MS Gothic
      And font.name_east_asia is MS Gothic
