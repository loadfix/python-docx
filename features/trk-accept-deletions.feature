Feature: Accept tracked deletions
  In order to finalise a document by discarding runs the reviewer deleted
  As a developer using python-docx
  I need Document.accept_all_changes and TrackedChange.accept to remove every
  w:del wrapper and its deleted text while leaving live content intact


  Scenario: accept_all_changes removes every w:del and w:delText
    Given the trk-accept-del document
     When I call document.accept_all_changes
     Then the accept-changes count is 7
      And the document has no w:del elements
      And the document has no w:delText elements
      And the document has no w:cellDel elements


  Scenario: accept_all_changes leaves the surrounding live text intact
    Given the trk-accept-del document
     When I call document.accept_all_changes
     Then paragraph 1 text is "Keep A  end A."
      And paragraph 5 text is "survivor"


  Scenario: per-change accept on a single deletion drops that w:del only
    Given the trk-accept-del document
     When I accept the only tracked change on paragraph 1
     Then paragraph 1 has no tracked changes
      And paragraph 1 text is "Keep A  end A."
      And paragraph 2 still has 1 tracked change
      And the document has 4 w:del elements


  Scenario: accepting a deletion that spans multiple runs collapses all of them
    Given the trk-accept-del document
     When I accept the only tracked change on paragraph 2
     Then paragraph 2 text is "Span  tail."
      And paragraph 2 has no w:del children
      And paragraph 2 has no w:delText descendants


  Scenario: accepting deletions beside an insertion leaves the insertion alone
    Given the trk-accept-del document
     When I accept every deletion-typed tracked change on paragraph 3
     Then paragraph 3 has 1 tracked change of type "insertion"
      And paragraph 3 has no w:del children
      And paragraph 3 still has a w:ins child


  Scenario: accepting a whole-paragraph deletion removes pPr/rPr/w:del too
    Given the trk-accept-del document
     When I call document.accept_all_changes
     Then paragraph 4 has no w:pPr/w:rPr/w:del marker
      And paragraph 4 has no w:del children
      And paragraph 4 text is ""


  Scenario: accepting a cellDel removes the whole cell from its row
    Given the trk-accept-del document
     When I call document.accept_all_changes
     Then the first table has 1 row with 1 cell
      And the first cell text is "kept cell"


  Scenario: accept_all_changes is idempotent
    Given the trk-accept-del document
     When I call document.accept_all_changes
      And I call document.accept_all_changes
     Then the accept-changes count is 0
      And the document has no w:del elements
