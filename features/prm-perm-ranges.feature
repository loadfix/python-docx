Feature: Mark portions of a document as editable by specific users or groups
  In order to allow structured editing while the rest of the document is locked
  As a developer using python-docx
  I need Paragraph.add_permission_range and Document.permission_ranges


  Scenario: Document.permission_ranges enumerates permStart elements in order
    Given a document having three permission ranges
     Then document.permission_ranges has length 3
      And permission_ranges[0].edit_group == "everyone"
      And permission_ranges[1].user == "alice@example.com"
      And permission_ranges[2].edit_group == "Authors"


  Scenario: Permission range ids are assigned sequentially from 0
    Given a document having three permission ranges
     Then permission_ranges have ids [0, 1, 2]


  Scenario: Adding a permission range wraps the paragraph in permStart/permEnd
    Given a fresh default document with one paragraph
     When I call paragraph.add_permission_range(edit_group="everyone")
     Then document.permission_ranges has length 1
      And permission_ranges[0].edit_group == "everyone"


  Scenario: Adding a user-restricted permission range
    Given a fresh default document with one paragraph
     When I call paragraph.add_permission_range(user="alice@example.com")
     Then permission_ranges[0].user == "alice@example.com"
      And permission_ranges[0].edit_group is None


  Scenario: Deleting a permission range removes both permStart and permEnd
    Given a document having three permission ranges
     When I call permission_ranges[0].delete()
     Then document.permission_ranges has length 2
