Feature: Get and set table row properties
  In order to format a table row to my requirements
  As a developer using python-docx
  I need a way to get and set the properties of a table row


  Scenario Outline: Get Row.grid_cols_after
    Given a table row ending with <count> empty grid columns
     Then row.grid_cols_after is <count>

    Examples: Row.grid_cols_after value cases
      | count |
      | 0     |
      | 1     |
      | 2     |


  Scenario Outline: Get Row.grid_cols_before
    Given a table row starting with <count> empty grid columns
     Then row.grid_cols_before is <count>

    Examples: Row.grid_cols_before value cases
      | count |
      | 0     |
      | 1     |
      | 3     |


  Scenario Outline: Get Row.height_rule
    Given a table row having height rule <state>
     Then row.height_rule is <value>

    Examples: Row.height_rule value cases
      | state               | value    |
      | no explicit setting | None     |
      | automatic           | AUTO     |
      | at least            | AT_LEAST |


  Scenario Outline: Set Row.height_rule
    Given a table row having height rule <state>
     When I assign <value> to row.height_rule
     Then row.height_rule is <value>

    Examples: Row.height_rule assignment cases
      | state               | value    |
      | no explicit setting | AUTO     |
      | automatic           | AT_LEAST |
      | at least            | None     |
      | no explicit setting | None     |


  Scenario Outline: Get Row.height
    Given a table row having height of <state>
     Then row.height is <value>

    Examples: Row.height value cases
      | state               | value   |
      | no explicit setting | None    |
      | 2 inches            | 1828800 |
      | 3 inches            | 2743200 |


  Scenario Outline: Set row height
    Given a table row having height of <state>
     When I assign <value> to row.height
     Then row.height is <value>

    Examples: Row.height assignment cases
      | state               | value   |
      | no explicit setting | 1828800 |
      | 2 inches            | 2743200 |
      | 3 inches            | None    |
      | no explicit setting | None    |


  Scenario: Row.allow_break_across_pages defaults to True
    Given a row in a freshly-created table
     Then row.allow_break_across_pages is True


  Scenario Outline: Set Row.allow_break_across_pages round-trips
    Given a row in a freshly-created table
     When I assign <value> to row.allow_break_across_pages
     Then row.allow_break_across_pages is <value>

    Examples: Allow-break-across-pages values
      | value |
      | True  |
      | False |


  Scenario: Row.is_header defaults to False
    Given a row in a freshly-created table
     Then row.is_header is False


  Scenario Outline: Set Row.is_header round-trips
    Given a row in a freshly-created table
     When I assign <value> to row.is_header
     Then row.is_header is <value>

    Examples: Is-header values
      | value |
      | True  |
      | False |
