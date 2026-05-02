Feature: Get and set per-cell margins
  In order to control the inner padding of a table cell
  As a developer using python-docx
  I need to read and write each edge of a cell's ``w:tcMar`` element


  Scenario Outline: Read Cell.margins.<edge> when set
    Given a cell having explicit margins on every edge
     Then cell.margins.<edge> is <value>

    Examples: Explicit cell-margin edge values
      | edge   | value |
      | top    | 45720 |
      | bottom | 45720 |
      | start  | 73025 |
      | end    | 73025 |


  Scenario Outline: Read Cell.margins.<edge> when unset
    Given a cell having no explicit margins
     Then cell.margins.<edge> is None

    Examples: Cell-margin edges with no directly-applied value
      | edge   |
      | top    |
      | bottom |
      | start  |
      | end    |


  Scenario Outline: Set Cell.margins.<edge> creates the edge on demand
    Given a cell having no explicit margins
     When I assign 114300 to cell.margins.<edge>
     Then cell.margins.<edge> is 114300

    Examples: Cell-margin edges being assigned
      | edge   |
      | top    |
      | bottom |
      | start  |
      | end    |


  Scenario: Assigning None to a cell-margin edge removes it
    Given a cell having explicit margins on every edge
     When I assign None to cell.margins.top
     Then cell.margins.top is None
      And cell.margins.bottom is 45720
      And cell.margins.start is 73025
      And cell.margins.end is 73025


  Scenario: Cell.set_margins writes only the edges I specify
    Given a cell having no explicit margins
     When I call cell.set_margins(top=114300, end=228600)
     Then cell.margins.top is 114300
      And cell.margins.bottom is None
      And cell.margins.start is None
      And cell.margins.end is 228600


  Scenario: Cell.remove_margins clears all edges
    Given a cell having explicit margins on every edge
     When I call cell.remove_margins()
     Then cell.margins.top is None
      And cell.margins.bottom is None
      And cell.margins.start is None
      And cell.margins.end is None
