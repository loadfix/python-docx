Feature: Get and set table style conditional flags
  In order to control which parts of a table style are applied
  As a developer using python-docx
  I need to read and write the flags on a table's ``w:tblLook`` element


  Scenario Outline: Read flag on a table with no explicit w:tblLook
    Given the tbl-banded table without any tblLook flags
     Then table.style_flags.<flag> is False

    Examples: Flags on a default tblLook
      | flag                  |
      | first_row             |
      | last_row              |
      | first_column          |
      | last_column           |
      | no_horizontal_banding |
      | no_vertical_banding   |


  Scenario: Read first_row flag when set by fixture
    Given the tbl-banded table with only first_row set
     Then table.style_flags.first_row is True
      And table.style_flags.first_column is False


  Scenario: Read no_horizontal_banding when explicitly off
    Given the tbl-banded table with banded rows active
     Then table.style_flags.first_row is True
      And table.style_flags.first_column is True
      And table.style_flags.no_horizontal_banding is False


  Scenario: Read no_horizontal_banding when explicitly on
    Given the tbl-banded table with banded rows suppressed
     Then table.style_flags.no_horizontal_banding is True


  Scenario Outline: Set a style flag creates the w:tblLook element on demand
    Given the tbl-banded table without any tblLook flags
     When I assign True to table.style_flags.<flag>
     Then table.style_flags.<flag> is True

    Examples: Flags that can be flipped on
      | flag                  |
      | first_row             |
      | last_row              |
      | first_column          |
      | last_column           |
      | no_horizontal_banding |
      | no_vertical_banding   |


  Scenario: Clear a style flag by writing False
    Given the tbl-banded table with only first_row set
     When I assign False to table.style_flags.first_row
     Then table.style_flags.first_row is False
