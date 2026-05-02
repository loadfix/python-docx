Feature: Get and set cell text direction
  In order to rotate text within a table cell
  As a developer using python-docx
  I need to read and write the ``w:textDirection`` child of ``w:tcPr``


  Scenario: Read Cell.text_direction when unset
    Given a table cell
     Then cell.text_direction is None


  Scenario Outline: Set and read Cell.text_direction
    Given a table cell
     When I assign <value> to cell.text_direction
     Then cell.text_direction is <value>

    Examples: Supported text-direction values
      | value                   |
      | WD_TEXT_DIRECTION.LR_TB |
      | WD_TEXT_DIRECTION.TB_RL |
      | WD_TEXT_DIRECTION.BT_LR |
      | WD_TEXT_DIRECTION.LR_TB_V |
      | WD_TEXT_DIRECTION.TB_RL_V |
      | WD_TEXT_DIRECTION.TB_LR_V |


  Scenario: Assigning None clears cell text direction
    Given a table cell
     When I assign WD_TEXT_DIRECTION.TB_RL to cell.text_direction
      And I assign None to cell.text_direction
     Then cell.text_direction is None
