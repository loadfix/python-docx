Feature: East Asian document grid
  In order to configure the line / character grid for East Asian layout
  As a developer using python-docx
  I need read/write access to Section.document_grid and related helpers


  Scenario: Default section exposes a DocumentGrid proxy
    Given a Section with default document grid as section
     Then section.document_grid is a DocumentGrid object
      And section.document_grid.type is None
      And section.document_grid.line_pitch is 360
      And section.document_grid.char_space is None


  Scenario: Fully-populated docGrid reads back
    Given a Section with fully populated document grid as section
     Then section.document_grid is a DocumentGrid object
      And section.document_grid.type is LINES_AND_CHARS
      And section.document_grid.line_pitch is 312
      And section.document_grid.char_space is 0


  Scenario: No docGrid element present
    Given a Section with no document grid as section
     Then section.document_grid is None


  Scenario: set_document_grid() preserves untouched attributes
    Given a Section with fully populated document grid as section
     When I call section.set_document_grid with line_pitch 480 only
     Then section.document_grid.line_pitch is 480
      And section.document_grid.type is LINES_AND_CHARS
      And section.document_grid.char_space is 0


  Scenario: set_document_grid() creates the element lazily
    Given a Section with no document grid as section
     When I call section.set_document_grid with type LINES and line_pitch 240
     Then section.document_grid is a DocumentGrid object
      And section.document_grid.type is LINES
      And section.document_grid.line_pitch is 240


  Scenario Outline: Individual attribute setters
    Given a Section with fully populated document grid as section
     When I assign <value> to section.document_grid.<attr>
     Then section.document_grid.<attr> is <reported>

    Examples: Direct setters
      | attr       | value        | reported     |
      | type       | SNAP_TO_CHARS| SNAP_TO_CHARS|
      | line_pitch | 420          | 420          |
      | char_space | 10           | 10           |


  Scenario: Remove document grid element
    Given a Section with fully populated document grid as section
     When I call section.remove_document_grid()
     Then section.document_grid is None


  Scenario: remove_document_grid() is a no-op when none present
    Given a Section with no document grid as section
     When I call section.remove_document_grid()
     Then section.document_grid is None
