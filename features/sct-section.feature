Feature: Access and change section properties
  In order to discover and modify document section behaviors
  As a developer using python-docx
  I need a way to get and set the properties of a section


  Scenario Outline: Section.different_first_page_header_footer getter
    Given a Section object <with-or-without> a distinct first-page header as section
     Then section.different_first_page_header_footer is <value>

    Examples: Section.different_first_page_header_footer states
      | with-or-without | value |
      | with            | True  |
      | without         | False |


  Scenario Outline: Section.different_first_page_header_footer setter
    Given a Section object <with-or-without> a distinct first-page header as section
     When I assign <value> to section.different_first_page_header_footer
     Then section.different_first_page_header_footer is <value>

    Examples: Section.different_first_page_header_footer assignment cases
      | with-or-without | value |
      | with            | True  |
      | with            | False |
      | without         | True  |
      | without         | False |


  Scenario: Section.even_page_footer
    Given a Section object as section
     Then section.even_page_footer is a _Footer object


  Scenario: Section.even_page_header
    Given a Section object as section
     Then section.even_page_header is a _Header object


  Scenario: Section.first_page_footer
    Given a Section object as section
     Then section.first_page_footer is a _Footer object


  Scenario: Section.first_page_header
    Given a Section object as section
     Then section.first_page_header is a _Header object


  Scenario: Section.footer
    Given a Section object as section
     Then section.footer is a _Footer object


  Scenario: Section.header
    Given a Section object as section
     Then section.header is a _Header object


  Scenario: Section.iter_inner_content()
    Given a Section object of a multi-section document as section
     Then section.iter_inner_content() produces the paragraphs and tables in section


  Scenario Outline: Get section start type
    Given a section having start type <start-type>
     Then the reported section start type is <start-type>

    Examples: Section start types
      | start-type |
      | CONTINUOUS |
      | NEW_COLUMN |
      | NEW_PAGE   |
      | EVEN_PAGE  |
      | ODD_PAGE   |


  Scenario Outline: Set section start type
    Given a section having start type <initial-start-type>
     When I set the section start type to <new-start-type>
     Then the reported section start type is <reported-start-type>

    Examples: Section start types
      | initial-start-type | new-start-type | reported-start-type |
      | CONTINUOUS         | NEW_PAGE       | NEW_PAGE            |
      | NEW_PAGE           | ODD_PAGE       | ODD_PAGE            |
      | NEW_COLUMN         | None           | NEW_PAGE            |


  Scenario: Get section page size
    Given a section having known page dimension
     Then the reported page width is 8.5 inches
      And the reported page height is 11 inches


  Scenario: Set section page size
    Given a section having known page dimension
     When I set the section page width to 11 inches
      And I set the section page height to 8.5 inches
     Then the reported page width is 11 inches
      And the reported page height is 8.5 inches


  Scenario Outline: Get section orientation
    Given a section known to have <orientation> orientation
     Then the reported page orientation is <reported-orientation>

    Examples: Section page orientations
      | orientation | reported-orientation |
      | landscape   | WD_ORIENT.LANDSCAPE  |
      | portrait    | WD_ORIENT.PORTRAIT   |


  Scenario Outline: Set section orientation
    Given a section known to have <initial-orientation> orientation
     When I set the section orientation to <new-orientation>
     Then the reported page orientation is <reported-orientation>

    Examples: Section page orientations
      | initial-orientation | new-orientation      |  reported-orientation |
      | portrait            | WD_ORIENT.LANDSCAPE  |  WD_ORIENT.LANDSCAPE  |
      | landscape           | WD_ORIENT.PORTRAIT   |  WD_ORIENT.PORTRAIT   |
      | landscape           | None                 |  WD_ORIENT.PORTRAIT   |


  Scenario: Get section page margins
    Given a section having known page margins
     Then the reported left margin is 1.0 inches
      And the reported right margin is 1.25 inches
      And the reported top margin is 1.5 inches
      And the reported bottom margin is 1.75 inches
      And the reported gutter margin is 0.25 inches
      And the reported header margin is 0.5 inches
      And the reported footer margin is 0.75 inches


  Scenario Outline: Set section page margins
    Given a section having known page margins
     When I set the <margin-type> margin to <length> inches
     Then the reported <margin-type> margin is <length> inches

    Examples: Section margin settings
      | margin-type | length |
      | left        |  1.0   |
      | right       |  1.25  |
      | top         |  0.75  |
      | bottom      |  1.5   |
      | header      |  0.25  |
      | footer      |  0.5   |
      | gutter      |  0.25  |


  # -- multi-column layout --------------------------------------------------

  Scenario: Default section reports a single column
    Given a Section with a single column as section
     Then section.columns is a SectionColumns object
      And section.columns.count is 1
      And section.columns.equal_width is True
      And section.columns.space is None
      And len(section.columns) is 0


  Scenario: Read equal-width multi-column settings
    Given a Section with three equal columns as section
     Then section.columns.count is 3
      And section.columns.equal_width is True
      And section.columns.space is 18 pt
      And len(section.columns) is 0


  Scenario: Read unequal-width multi-column settings
    Given a Section with two unequal columns as section
     Then section.columns.count is 2
      And section.columns.equal_width is False
      And len(section.columns) is 2
      And section.columns[0].width is 2.5 inches
      And section.columns[0].space is 0.5 inches
      And section.columns[1].width is 4.0 inches


  Scenario: Iterate over the columns sequence
    Given a Section with two unequal columns as section
     Then iterating section.columns yields 2 Column objects


  Scenario Outline: Set section column count and spacing
    Given a Section with two equal columns as section
     When I assign <count> to section.columns.count
      And I assign <space> to section.columns.space in pt
     Then section.columns.count is <count>
      And section.columns.space is <space> pt

    Examples: Column count / space combinations
      | count | space |
      |     2 |    12 |
      |     3 |    24 |
      |     4 |    18 |


  Scenario: Toggle equal_width
    Given a Section with two equal columns as section
     When I assign False to section.columns.equal_width
     Then section.columns.equal_width is False


  Scenario: Update individual Column width
    Given a Section with two unequal columns as section
     When I assign 3.0 inches to section.columns[0].width
      And I assign 0.75 inches to section.columns[1].space
     Then section.columns[0].width is 3.0 inches
      And section.columns[1].space is 0.75 inches
