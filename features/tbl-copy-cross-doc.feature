Feature: Copy a table across documents
  In order to reuse a table defined in another document
  As a developer using python-docx
  I need a way to deep-copy a table element between two Documents, rewiring its
  image and style references into the destination document's package.

  Scenario: Copy a 2x2 table with a styled header and an embedded PNG
     Given a source document containing a 2x2 table with a styled header and an embedded PNG
       And an empty destination document
      When I call add_table_copy on the destination with the source table
      Then the destination contains one table with the copied cell text
       And the destination contains at least one image part
       And the copied table's embedded image reference resolves in the destination
