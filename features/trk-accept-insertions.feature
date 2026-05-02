Feature: Accept tracked insertions
  In order to resolve inserted content a reviewer proposed via track changes
  As a developer using python-docx
  I need to accept insertions at both the document and individual-change level


  Scenario: accept_all_changes removes every w:ins wrapper
    Given the trk-accept-ins document
     When I call document.accept_all_changes()
     Then the document body has 0 w:ins elements
      And the document body has 0 w:del elements
      And the document body has 0 w:cellIns elements


  Scenario: Accepting an insertion leaves the inserted text as live content
    Given the trk-accept-ins document
     When I call document.accept_all_changes()
     Then paragraph 1 text == "The quick nimble fox jumps."
      And paragraph 1 has no tracked changes


  Scenario: Accepting an insertion that spans multiple runs keeps every run
    Given the trk-accept-ins document
     When I call document.accept_all_changes()
     Then paragraph 2 text == "Red fish blue fish"
      And paragraph 2 has 3 direct w:r children


  Scenario: Per-change accept resolves only the targeted insertion
    Given the trk-accept-ins document
     When I accept the first tracked change of paragraph 3
     Then paragraph 3 has 1 tracked change remaining
      And paragraph 3 tracked_change[0].author == "Carol"
      And paragraph 3 text == "Alpha beta gamma"


  Scenario: Per-change accept on a multi-run insertion unwraps every run
    Given the trk-accept-ins document
     When I accept the first tracked change of paragraph 2
     Then paragraph 2 has 0 tracked changes remaining
      And paragraph 2 text == "Red fish blue fish"


  Scenario: accept_all_changes cleans up a paragraph that also has a deletion
    Given the trk-accept-ins document
     When I call document.accept_all_changes()
     Then paragraph 4 text == "Hello bright world."
      And paragraph 4 has no tracked changes


  Scenario: Accepting a cellIns keeps the cell but removes the wrapper
    Given the trk-accept-ins document
     When I call document.accept_all_changes()
     Then table 0 cell (0, 1) text == "inserted"
      And table 0 cell (0, 1).is_tracked_insertion is False


  Scenario: accept_all_changes returns the number of revisions resolved
    Given the trk-accept-ins document
     When I call document.accept_all_changes()
     Then the accept_all_changes return value is 7
