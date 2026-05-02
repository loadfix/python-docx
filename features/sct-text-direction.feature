Feature: Section text direction and right-to-left flow
  In order to support vertical East Asian layouts and right-to-left languages
  As a developer using python-docx
  I need read/write access to Section.text_direction and Section.right_to_left


  Scenario: Default section has no textDirection and LTR
    Given a Section with default text direction as section
     Then section.text_direction is None
      And section.right_to_left is False


  Scenario: RTL and TB_RL read back
    Given a Section with TB_RL vertical RTL as section
     Then section.text_direction is TB_RL
      And section.right_to_left is True


  Scenario: BT_LR vertical with LTR reads back
    Given a Section with BT_LR vertical LTR as section
     Then section.text_direction is BT_LR
      And section.right_to_left is False


  Scenario Outline: Assign text_direction
    Given a Section with default text direction as section
     When I assign <value> to section.text_direction
     Then section.text_direction is <reported>

    Examples: Text direction assignments
      | value  | reported |
      | TB_RL  | TB_RL    |
      | BT_LR  | BT_LR    |
      | LR_TB  | LR_TB    |


  Scenario: Clearing text_direction removes the element
    Given a Section with TB_RL vertical RTL as section
     When I assign None to section.text_direction
     Then section.text_direction is None


  Scenario Outline: Assign right_to_left
    Given a Section with default text direction as section
     When I assign <value> to section.right_to_left
     Then section.right_to_left is <reported>

    Examples: RTL assignments
      | value | reported |
      | True  | True     |
      | False | False    |


  Scenario: Assigning False removes bidi element
    Given a Section with TB_RL vertical RTL as section
     When I assign False to section.right_to_left
     Then section.right_to_left is False


  Scenario: Assigning None treats as False
    Given a Section with TB_RL vertical RTL as section
     When I assign None to section.right_to_left
     Then section.right_to_left is False
