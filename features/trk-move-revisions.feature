Feature: MoveRevision for paired w:moveFrom and w:moveTo
  In order to follow move revisions as a connected pair rather than isolated wrappers
  As a developer using python-docx
  I need MoveRevision proxies exposing the shared w:name and a .peer lookup


  Scenario: Paragraph.tracked_changes wraps w:moveFrom as MoveRevision
    Given the trk-move document
     When I select the tracked_changes of paragraph 1
     Then the first tracked change is a MoveRevision
      And tracked_change[0].type == "move_from"
      And tracked_change[0].name == "pair1"


  Scenario: Paragraph.tracked_changes wraps w:moveTo as MoveRevision
    Given the trk-move document
     When I select the tracked_changes of paragraph 2
     Then the first tracked change is a MoveRevision
      And tracked_change[0].type == "move_to"
      And tracked_change[0].name == "pair1"


  Scenario: MoveRevision.peer resolves the destination from the source
    Given the trk-move document
     When I select the tracked_changes of paragraph 1
     Then the peer of the first move revision has type "move_to"
      And the peer of the first move revision has name "pair1"


  Scenario: MoveRevision.peer resolves the source from the destination
    Given the trk-move document
     When I select the tracked_changes of paragraph 2
     Then the peer of the first move revision has type "move_from"
      And the peer of the first move revision has name "pair1"
