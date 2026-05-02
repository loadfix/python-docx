Feature: Get and set style properties
  In order to adjust a style to suit my needs
  As a developer using python-docx
  I need a set of read/write style properties


  Scenario Outline: Get base style
    Given a style based on <base-style>
     Then style.base_style is <value>

    Examples: Base style values
      | base-style | value            |
      | no style   | None             |
      | Normal     | styles['Normal'] |


  Scenario Outline: Set base style
    Given a style based on <base-style>
     When I assign <assigned-value> to style.base_style
     Then style.base_style is <value>

    Examples: Base style values
      | base-style | assigned-value   | value            |
      | no style   | styles['Normal'] | styles['Normal'] |
      | Normal     | styles['Base']   | styles['Base']   |
      | Base       | None             | None             |


  Scenario Outline: Get hidden value
    Given a style having hidden set <setting>
     Then style.hidden is <value>

    Examples: Style hidden values
      | setting    | value |
      | on         | True  |
      | off        | False |
      | no setting | False |


  Scenario Outline: Set hidden value
    Given a style having hidden set <setting>
     When I assign <new-value> to style.hidden
     Then style.hidden is <value>

    Examples: Style hidden values
      | setting    | new-value | value |
      | no setting | True      | True  |
      | on         | False     | False |


  Scenario Outline: Get locked value
    Given a style having locked set <setting>
     Then style.locked is <value>

    Examples: Style locked values
      | setting    | value |
      | on         | True  |
      | off        | False |
      | no setting | False |


  Scenario Outline: Set locked value
    Given a style having locked set <setting>
     When I assign <new-value> to style.locked
     Then style.locked is <value>

    Examples: Style locked values
      | setting    | new-value | value |
      | no setting | True      | True  |
      | on         | False     | False |


  Scenario: Get name
    Given a style having a known name
     Then style.name is the known name


  Scenario: Set name
    Given a style having a known name
     When I assign a new name to the style
     Then style.name is the new name


  Scenario Outline: Get next paragraph style
    Given a style having next paragraph style set to <setting>
     Then style.next_paragraph_style is <value>

    Examples: Style next paragraph style values
      | setting    | value      |
      | no setting | Base       |
      | Sub Normal | Sub Normal |
      | Foobar     | Sub Normal |


  Scenario Outline: Set next paragraph style
    Given a style having next paragraph style set to <setting>
     When I assign <new-value> to style.next_paragraph_style
     Then style.next_paragraph_style is <value>

    Examples: Results of assignment to .next_paragraph_style
      | setting    | new-value  | value      |
      | no setting | Citation   | Citation   |
      | Sub Normal | Base       | Base       |
      | Base       | None       | Foo        |


  Scenario Outline: Get style display sort order
    Given a style having priority of <setting>
     Then style.priority is <value>

    Examples: style.priority values
      | setting    | value |
      | no setting | None  |
      | 42         | 42    |


  Scenario Outline: Set style display sort order
    Given a style having priority of <setting>
     When I assign <new-value> to style.priority
     Then style.priority is <value>

    Examples: Style priority values
      | setting    | new-value | value |
      | no setting | 42        | 42    |
      | 42         | 24        | 24    |
      | 42         | None      | None  |


  Scenario Outline: Get quick-style value
    Given a style having quick-style set <setting>
     Then style.quick_style is <value>

    Examples: Style quick-style values
      | setting    | value |
      | on         | True  |
      | off        | False |
      | no setting | False |


  Scenario Outline: Set quick-style value
    Given a style having quick-style set <setting>
     When I assign <new-value> to style.quick_style
     Then style.quick_style is <value>

    Examples: Style quick_style values
      | setting    | new-value | value |
      | no setting | True      | True  |
      | on         | False     | False |


  Scenario: Get style id
    Given a style having a known style id
     Then style.style_id is the known style id


  Scenario: Set style id
    Given a style having a known style id
     When I assign a new value to style.style_id
     Then style.style_id is the new style id


  Scenario: Get style type
    Given a style having a known type
     Then style.type is the known type


  Scenario Outline: Get unhide-when-used value
    Given a style having unhide-when-used set <setting>
     Then style.unhide_when_used is <value>

    Examples: Style unhide-when-used values
      | setting    | value |
      | on         | True  |
      | off        | False |
      | no setting | False |


  Scenario Outline: Set unhide-when-used value
    Given a style having unhide-when-used set <setting>
     When I assign <new-value> to style.unhide_when_used
     Then style.unhide_when_used is <value>

    Examples: Style unhide_when_used values
      | setting    | new-value | value |
      | no setting | True      | True  |
      | on         | False     | False |


  Scenario Outline: Get link_style
    Given a linked/next/redefined styles document
     Then styles[<name>].link_style is <expected>

    Examples: Style.link_style values
      | name       | expected  |
      | "Body"     | "BodyChar"|
      | "BodyChar" | "Body"    |
      | "Solo"     | None      |


  Scenario: Set link_style by style object
    Given a linked/next/redefined styles document
     When I assign styles["Body"] to styles["Solo"].link_style
     Then styles["Solo"].link_style is "Body"


  Scenario: Set link_style by style id string
    Given a linked/next/redefined styles document
     When I assign "Body" to styles["Solo"].link_style
     Then styles["Solo"].link_style is "Body"


  Scenario: Clear link_style
    Given a linked/next/redefined styles document
     When I assign None to styles["Body"].link_style
     Then styles["Body"].link_style is None


  Scenario Outline: Get next_style
    Given a linked/next/redefined styles document
     Then styles[<name>].next_style is <expected>

    Examples: Style.next_style values
      | name    | expected |
      | "Intro" | "Body"   |
      | "Solo"  | None     |


  Scenario: Set next_style by style object
    Given a linked/next/redefined styles document
     When I assign styles["Body"] to styles["Solo"].next_style
     Then styles["Solo"].next_style is "Body"


  Scenario: Clear next_style
    Given a linked/next/redefined styles document
     When I assign None to styles["Intro"].next_style
     Then styles["Intro"].next_style is None


  Scenario Outline: Get is_redefined
    Given a linked/next/redefined styles document
     Then styles[<name>].is_redefined is <value>

    Examples: Style.is_redefined values
      | name    | value |
      | "Intro" | True  |
      | "Body"  | False |
      | "Solo"  | False |
