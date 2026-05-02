Feature: Insert a paragraph or table at an arbitrary position
  In order to add new block-level content anywhere in the document
  As a developer using python-docx
  I need position-aware insert helpers on Paragraph and Table


  Scenario: Paragraph.insert_paragraph_after inserts directly after a reference paragraph
    Given a document containing five paragraphs
     When I insert a paragraph after the third paragraph
     Then the document contains six paragraphs
      And the fourth paragraph text is "inserted-after"


  Scenario: Paragraph.insert_paragraph_before inserts directly before a reference paragraph
    Given a document containing five paragraphs
     When I insert a paragraph before the fourth paragraph
     Then the document contains six paragraphs
      And the fourth paragraph text is "inserted-before"


  Scenario: Paragraph.insert_table_after inserts a table directly after the reference paragraph
    Given a document containing five paragraphs
     When I insert a 2x2 table after the third paragraph
     Then the document contains one table
      And the inserted table has two rows and two columns


  Scenario: Paragraph.insert_table_before inserts a table directly before the reference paragraph
    Given a document containing five paragraphs
     When I insert a 2x2 table before the fourth paragraph
     Then the document contains one table
      And the inserted table has two rows and two columns


  Scenario: Table.insert_paragraph_after inserts a paragraph directly after a reference table
    Given a document containing three tables
     When I insert a paragraph after the second table
     Then the paragraph after the second table has text "after-table"


  Scenario: Table.insert_paragraph_before inserts a paragraph directly before a reference table
    Given a document containing three tables
     When I insert a paragraph before the second table
     Then the paragraph before the second table has text "before-table"


  Scenario: Table.insert_table_after inserts a new table directly after a reference table
    Given a document containing three tables
     When I insert a 2x2 table after the second table
     Then the document contains four tables
