Feature: Reject tracked changes
  In order to discard a reviewer's proposed edits
  As a developer using python-docx
  I need Document.reject_all_changes() and TrackedChange.reject() to unwind
  insertions, deletions, moves, and cell-level revisions


  Scenario: reject_all_changes on a document with only insertions removes the inserted content
    Given the trk-reject document with only insertions
     When I call document.reject_all_changes()
     Then the reject count is 2
      And the document has no w:ins elements
      And paragraph 1 text equals "The quick  fox jumps."
      And paragraph 2 text equals "Hello."


  Scenario: reject_all_changes on a document with only deletions restores the deleted content
    Given the trk-reject document with only deletions
     When I call document.reject_all_changes()
     Then the reject count is 2
      And the document has no w:del elements
      And the document has no w:delText elements
      And paragraph 1 text equals "The quick brown fox jumps."
      And paragraph 3 text equals "This whole paragraph is being deleted."


  Scenario: reject_all_changes on a mixed insertions-plus-deletions document
    Given the trk-reject document
     When I call document.reject_all_changes()
     Then the reject count is 8
      And the document has no w:ins elements
      And the document has no w:del elements
      And the document has no w:delText elements
      And paragraph 1 text equals "The quick brown fox jumps."
      And paragraph 2 text equals "Hello."


  Scenario: Per-change reject a single insertion leaves the sibling deletion untouched
    Given the trk-reject document
     When I reject the insertion in paragraph 1
     Then paragraph 1 has 1 tracked change remaining
      And the remaining tracked change in paragraph 1 is a deletion
      And paragraph 1 text equals "The quick  fox jumps."


  Scenario: Per-change reject a single deletion leaves the sibling insertion untouched
    Given the trk-reject document
     When I reject the deletion in paragraph 1
     Then paragraph 1 has 1 tracked change remaining
      And the remaining tracked change in paragraph 1 is an insertion
      And paragraph 1 revision_marks_text equals "The quick brown[+nimble+] fox jumps."


  Scenario: Rejecting a multi-run insertion removes every inserted run
    Given the trk-reject document
     When I reject every tracked change in paragraph 2
     Then paragraph 2 has 0 tracked changes remaining
      And paragraph 2 text equals "Hello."


  Scenario: Rejecting a deletion that wrapped an entire paragraph restores every fragment
    Given the trk-reject document
     When I reject every tracked change in paragraph 3
     Then paragraph 3 has 0 tracked changes remaining
      And paragraph 3 text equals "This whole paragraph is being deleted."


  Scenario: Rejecting a cellIns removes the inserted cell
    Given the trk-reject document
     When I call document.reject_all_changes()
     Then row 0 of table 0 has 1 cell
      And the document has no w:cellIns elements


  Scenario: Rejecting a cellDel keeps the previously-deleted cell
    Given the trk-reject document
     When I call document.reject_all_changes()
     Then row 1 of table 0 has 2 cells
      And the document has no w:cellDel elements


  Scenario: Rejecting a move revision restores the source and removes the destination
    Given the trk-reject document
     When I call document.reject_all_changes()
     Then the document has no w:moveFrom elements
      And the document has no w:moveTo elements
      And paragraph 5 text equals "Source: moved text (was here)."
      And paragraph 6 text equals "Destination:  (now here)."
