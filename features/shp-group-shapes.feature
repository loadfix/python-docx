Feature: Read DrawingML group shapes
  In order to inspect grouped shapes authored in Word
  As a developer using python-docx
  I need read access to ``wpg:grpSp`` group shapes and their children


  Scenario: Detect a top-level group shape on a drawing
    Given a document known to contain a DrawingML group shape
     Then the grouped drawing reports is_group is True
      And the grouped drawing type is GROUP


  Scenario: Enumerate top-level children of a group shape
    Given a document known to contain a DrawingML group shape
     Then the outer group has name "Outer Group"
      And the outer group has 3 top-level children
      And the first child is a WordprocessingShape of type RECTANGLE
      And the second child is a WordprocessingShape of type OVAL
      And the third child is a nested GroupShape


  Scenario: Traverse into a nested group shape
    Given a document known to contain a DrawingML group shape
     Then the nested group has name "Inner Group"
      And the nested group has 1 top-level child
      And the nested group's first child is a WordprocessingShape of type ARROW_RIGHT


  Scenario: Access text inside a grouped shape
    Given a document known to contain a DrawingML group shape
     Then the first shape inside the group has text "Alpha"
