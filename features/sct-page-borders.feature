Feature: Section page borders
  In order to draw decorative borders around printed pages
  As a developer using python-docx
  I need properties and methods to read, set, and remove page borders


  Scenario: PageBorders proxy is always available
    Given a Section with no page borders as section
     Then section.page_borders is a PageBorders object
      And section.page_borders.top is a PageBorder object
      And section.page_borders.top.style is None
      And section.page_borders.top.width is None
      And section.page_borders.top.color is None
      And section.page_borders.top.space is None


  Scenario: PageBorders.display and offset_from default to None
    Given a Section with no page borders as section
     Then section.page_borders.display is None
      And section.page_borders.offset_from is None


  Scenario Outline: Read pre-populated edge values
    Given a Section with all four borders set as section
     Then section.page_borders.<side>.style is SINGLE
      And section.page_borders.<side>.width is 1 pt
      And section.page_borders.<side>.color is FF0000
      And section.page_borders.<side>.space is 24 pt

    Examples: Each edge
      | side   |
      | top    |
      | bottom |
      | left   |
      | right  |


  Scenario: Read display and offset_from attributes
    Given a Section with all four borders set as section
     Then section.page_borders.display is ALL_PAGES
      And section.page_borders.offset_from is PAGE


  Scenario Outline: set_page_border() creates edge elements lazily
    Given a Section with no page borders as section
     When I call section.set_page_border on <side> with SINGLE 1pt black 12pt
     Then section.page_borders.<side>.style is SINGLE
      And section.page_borders.<side>.width is 1 pt
      And section.page_borders.<side>.space is 12 pt

    Examples: Each edge
      | side   |
      | top    |
      | bottom |
      | left   |
      | right  |


  Scenario: set_page_border() leaves untouched attributes alone
    Given a Section with all four borders set as section
     When I call section.set_page_border on top with space 6 pt
     Then section.page_borders.top.style is SINGLE
      And section.page_borders.top.width is 1 pt
      And section.page_borders.top.space is 6 pt


  Scenario: set_page_border() rejects invalid side
    Given a Section with no page borders as section
     Then calling section.set_page_border with side "inner" raises ValueError


  Scenario Outline: Clear individual border attributes
    Given a Section with all four borders set as section
     When I clear section.page_borders.top.<attr>
     Then section.page_borders.top.<attr> is None

    Examples: Clearable attributes
      | attr  |
      | style |
      | width |
      | color |
      | space |


  Scenario: Mutate display and offset_from
    Given a Section with all four borders set as section
     When I assign FIRST_PAGE to section.page_borders.display
      And I assign TEXT to section.page_borders.offset_from
     Then section.page_borders.display is FIRST_PAGE
      And section.page_borders.offset_from is TEXT


  Scenario: Remove all page borders
    Given a Section with all four borders set as section
     When I call section.remove_page_borders()
     Then section.page_borders.top.style is None
      And section.page_borders.display is None
      And section.page_borders.offset_from is None


  Scenario: remove_page_borders() is a no-op when none present
    Given a Section with no page borders as section
     When I call section.remove_page_borders()
     Then section.page_borders.top.style is None


  Scenario: Top-only border in mixed-state section
    Given a Section with only a top border set as section
     Then section.page_borders.top.style is THICK
      And section.page_borders.top.width is 3 pt
      And section.page_borders.top.color is 0000FF
      And section.page_borders.bottom.style is None
      And section.page_borders.left.style is None
      And section.page_borders.right.style is None
