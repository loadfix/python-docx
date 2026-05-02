Feature: Delete a run
  In order to remove a single run from inside a paragraph
  As a developer using python-docx
  I need a way to delete an existing run in place


  Scenario: Run.delete() removes the run from its paragraph
    Given a paragraph with three runs
     When I delete the second run
     Then the paragraph contains two runs
      And the paragraph run text sequence is "alpha |charlie"


  Scenario: Run.delete() leaves sibling runs intact
    Given a paragraph with three runs
     When I delete the first run
     Then the paragraph contains two runs
      And the paragraph text is "bravo charlie"


  Scenario: Run.delete() on a detached run is a no-op
    Given a detached run
     When I delete the detached run
     Then no error is raised
