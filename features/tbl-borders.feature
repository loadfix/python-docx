Feature: Get and set table and cell borders
  In order to control table and cell border appearance
  As a developer using python-docx
  I need a way to read and write each border edge of a table or cell


  Scenario Outline: Get Table.borders.<edge> style
    Given a table having borders on every edge
     Then table.borders.<edge>.style is WD_BORDER_STYLE.SINGLE
      And table.borders.<edge>.width is 6350
      And table.borders.<edge>.color is 000000

    Examples: Table border edges
      | edge      |
      | top       |
      | bottom    |
      | left      |
      | right     |
      | inside_h  |
      | inside_v  |


  Scenario Outline: Read Table.borders.<edge> when unset
    Given a table having no explicit borders
     Then table.borders.<edge>.style is None
      And table.borders.<edge>.width is None
      And table.borders.<edge>.color is None

    Examples: Edges with no directly-applied border
      | edge      |
      | top       |
      | bottom    |
      | left      |
      | right     |
      | inside_h  |
      | inside_v  |


  Scenario: Set Table.borders.top creates the border on demand
    Given a table having no explicit borders
     When I assign WD_BORDER_STYLE.SINGLE, 12700, 4F81BD to table.borders.top
     Then table.borders.top.style is WD_BORDER_STYLE.SINGLE
      And table.borders.top.width is 12700
      And table.borders.top.color is 4F81BD


  Scenario: Table.set_borders sets the specified edges to SINGLE and clears others
    Given a table having no explicit borders
     When I call table.set_borders(top=True, bottom=True, inside_h=True)
     Then table.borders.top.style is WD_BORDER_STYLE.SINGLE
      And table.borders.bottom.style is WD_BORDER_STYLE.SINGLE
      And table.borders.inside_h.style is WD_BORDER_STYLE.SINGLE
      And table.borders.left.style is WD_BORDER_STYLE.NONE
      And table.borders.right.style is WD_BORDER_STYLE.NONE
      And table.borders.inside_v.style is WD_BORDER_STYLE.NONE


  Scenario: Read Cell.borders when set
    Given a cell having a THICK left border
     Then cell.borders.left.style is WD_BORDER_STYLE.THICK
      And cell.borders.left.width is 12700
      And cell.borders.left.color is FF0000


  Scenario Outline: Read Cell.borders.<edge> when unset
    Given a cell having no explicit borders
     Then cell.borders.<edge>.style is None

    Examples: Cell edges with no directly-applied border
      | edge   |
      | top    |
      | bottom |
      | left   |
      | right  |


  Scenario: Set Cell.borders.right creates the border on demand
    Given a cell having no explicit borders
     When I assign WD_BORDER_STYLE.DOUBLE, 6350, 00FF00 to cell.borders.right
     Then cell.borders.right.style is WD_BORDER_STYLE.DOUBLE
      And cell.borders.right.width is 6350
      And cell.borders.right.color is 00FF00


  Scenario: Clear a border by assigning None to style
    Given a cell having a THICK left border
     When I assign None to cell.borders.left.style
     Then cell.borders.left.style is None
