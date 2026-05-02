Feature: Delete a table
  In order to remove a table from the middle of a document
  As a developer using python-docx
  I need a way to delete an existing table in place


  Scenario: Table.delete() removes the table from the body
    Given a document containing three tables
     When I delete the second table
     Then the document contains two tables
      And the remaining tables contain text "T0 r0c0" and "T2 r0c0"


  Scenario: Table.delete() preserves surrounding paragraphs
    Given a document containing three tables
     When I delete the second table
     Then the document paragraph text contains "Before table 1"
      And the document paragraph text contains "After table 3"


  Scenario: Table.delete() on a detached table is a no-op
    Given a detached table
     When I delete the detached table
     Then no error is raised
