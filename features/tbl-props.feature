Feature: Get and set table properties
  In order to format a table to my requirements
  As a developer using python-docx
  I need a way to get and set a table's properties


  Scenario Outline: Get table alignment
    Given a table having <alignment> alignment
     Then table.alignment is <value>

    Examples: table alignment settings
      | alignment | value                     |
      | inherited | None                      |
      | left      | WD_TABLE_ALIGNMENT.LEFT   |
      | right     | WD_TABLE_ALIGNMENT.RIGHT  |
      | center    | WD_TABLE_ALIGNMENT.CENTER |


  Scenario Outline: Set table alignment
    Given a table having <alignment> alignment
     When I assign <value> to table.alignment
     Then table.alignment is <value>

    Examples: results of assignment to table.alignment
      | alignment | value                     |
      | inherited | WD_TABLE_ALIGNMENT.LEFT   |
      | left      | WD_TABLE_ALIGNMENT.RIGHT  |
      | right     | WD_TABLE_ALIGNMENT.CENTER |
      | center    | None                      |


  Scenario Outline: Get autofit layout setting
    Given a table having an autofit layout of <autofit-setting>
     Then the reported autofit setting is <reported-autofit>

    Examples: table autofit settings
      | autofit-setting     | reported-autofit |
      | no explicit setting | autofit          |
      | autofit             | autofit          |
      | fixed               | fixed            |


  Scenario Outline: Set autofit layout setting
    Given a table having an autofit layout of <autofit-setting>
     When I set the table autofit to <new-setting>
     Then the reported autofit setting is <reported-autofit>

    Examples: table column width values
      | autofit-setting     | new-setting | reported-autofit |
      | no explicit setting | autofit     | autofit          |
      | no explicit setting | fixed       | fixed            |
      | fixed               | autofit     | autofit          |
      | autofit             | autofit     | autofit          |
      | fixed               | fixed       | fixed            |
      | autofit             | fixed       | fixed            |


  Scenario Outline: Get table direction
    Given a table having table direction set <setting>
     Then table.table_direction is <value>

    Examples: Table on/off property values
      | setting       | value |
      | to inherit    | None  |
      | right-to-left | RTL   |
      | left-to-right | LTR   |


  Scenario Outline: Set table direction
    Given a table having table direction set <setting>
     When I assign <new-value> to table.table_direction
     Then table.table_direction is <value>

    Examples: Results of assignment to Table.table_direction
      | setting       | new-value | value |
      | to inherit    | RTL       | RTL   |
      | right-to-left | LTR       | LTR   |
      | left-to-right | None      | None  |


  Scenario Outline: Set Table.autofit_behavior round-trips
    Given a freshly-created table
     When I assign <value> to table.autofit_behavior
     Then table.autofit_behavior is <value>

    Examples: Autofit-behavior enum values
      | value                             |
      | WD_TABLE_AUTOFIT.FIXED_WIDTH      |
      | WD_TABLE_AUTOFIT.AUTOFIT_TO_WINDOW |
      | WD_TABLE_AUTOFIT.AUTOFIT_TO_CONTENTS |


  Scenario: Setting FIXED_WIDTH disables allow_autofit
    Given a freshly-created table
     When I assign WD_TABLE_AUTOFIT.FIXED_WIDTH to table.autofit_behavior
     Then table.allow_autofit is False
      And table.autofit is False


  Scenario: Setting AUTOFIT_TO_WINDOW enables allow_autofit
    Given a freshly-created table
     When I assign WD_TABLE_AUTOFIT.FIXED_WIDTH to table.autofit_behavior
      And I assign WD_TABLE_AUTOFIT.AUTOFIT_TO_WINDOW to table.autofit_behavior
     Then table.allow_autofit is True


  Scenario: Table.preferred_width defaults to None
    Given a freshly-created table
     Then table.preferred_width is None


  Scenario Outline: Set Table.preferred_width round-trips
    Given a freshly-created table
     When I assign <value> to table.preferred_width
     Then table.preferred_width is <value>

    Examples: Preferred-width values in EMU
      | value   |
      | 3657600 |
      | None    |


  Scenario Outline: Set Table.allow_autofit round-trips
    Given a freshly-created table
     When I assign <value> to table.allow_autofit
     Then table.allow_autofit is <value>

    Examples: allow_autofit values
      | value |
      | True  |
      | False |
