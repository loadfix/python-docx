Feature: Section line numbering
  In order to display line numbers in the margin of a section
  As a developer using python-docx
  I need properties and methods to read, set, and remove line numbering


  Scenario: Section with no line numbering reads as None
    Given a Section with no line numbering as section
     Then section.line_numbering is None


  Scenario: Read a fully-populated lnNumType
    Given a Section with fully populated line numbering as section
     Then section.line_numbering is a LineNumbering object
      And section.line_numbering.count_by is 1
      And section.line_numbering.start is 1
      And section.line_numbering.distance is 20 pt
      And section.line_numbering.restart is NEW_PAGE


  Scenario: Read a partially-populated lnNumType
    Given a Section with count_by only line numbering as section
     Then section.line_numbering.count_by is 5
      And section.line_numbering.start is None
      And section.line_numbering.distance is None
      And section.line_numbering.restart is None


  Scenario: set_line_numbering() creates the element lazily
    Given a Section with no line numbering as section
     When I call section.set_line_numbering with count_by 10 and restart CONTINUOUS
     Then section.line_numbering is a LineNumbering object
      And section.line_numbering.count_by is 10
      And section.line_numbering.restart is CONTINUOUS


  Scenario: set_line_numbering() preserves untouched attributes
    Given a Section with fully populated line numbering as section
     When I call section.set_line_numbering with count_by 2 only
     Then section.line_numbering.count_by is 2
      And section.line_numbering.start is 1
      And section.line_numbering.restart is NEW_PAGE


  Scenario Outline: Setter updates individual attributes
    Given a Section with fully populated line numbering as section
     When I assign <value> to section.line_numbering.<attr>
     Then section.line_numbering.<attr> is <reported>

    Examples: Individual setter cases
      | attr     | value        | reported     |
      | count_by | 7            | 7            |
      | start    | 100          | 100          |
      | restart  | NEW_SECTION  | NEW_SECTION  |


  Scenario: Remove line numbering
    Given a Section with fully populated line numbering as section
     When I call section.remove_line_numbering()
     Then section.line_numbering is None


  Scenario: remove_line_numbering() is a no-op when none present
    Given a Section with no line numbering as section
     When I call section.remove_line_numbering()
     Then section.line_numbering is None
