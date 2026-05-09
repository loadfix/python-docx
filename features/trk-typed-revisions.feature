Feature: Typed revision proxies and document-level revision accessors
  In order to inspect and resolve tracked changes by type
  As a developer using python-docx
  I need Insertion, Deletion, and Move subclasses plus `revisions`
  accessors on Document, Paragraph, and Run and matching
  accept_revisions() / reject_revisions() shortcut methods


  Scenario: Document.revisions lists every run-level revision in order
    Given a document with a tracked insertion, deletion, and move
     Then document.revisions has 4 entries
      And revisions[0] is an Insertion
      And revisions[1] is a Deletion
      And revisions[2] is a Move with type "move_from"
      And revisions[3] is a Move with type "move_to"


  Scenario: Paragraph.revisions mirrors Paragraph.tracked_changes
    Given a document with a tracked insertion and deletion in paragraph 1
     Then paragraph.revisions has 2 entries
      And paragraph.revisions[0] is a Deletion
      And paragraph.revisions[1] is an Insertion


  Scenario: Run.revisions reports the wrapping w:ins as an Insertion
    Given a run inside a tracked insertion
     Then run.revisions has 1 entry
      And run.revisions[0] is an Insertion


  Scenario: Run.revisions reports the w:rPrChange as a FormattingChange
    Given a run with a w:rPrChange on its w:rPr
     Then run.revisions has 1 entry
      And run.revisions[0] is a FormattingChange


  Scenario: accept_revisions() is an alias of accept_all_changes()
    Given a document with a tracked insertion and deletion
     When I call document.accept_revisions()
     Then every w:ins and w:del element has been removed
      And the returned count is 2


  Scenario: reject_revisions() is an alias of reject_all_changes()
    Given a document with a tracked insertion and deletion
     When I call document.reject_revisions()
     Then every w:ins and w:del element has been removed
      And the inserted content is gone
      And the deleted content is restored


  Scenario: Revision is an alias of TrackedChange for back-compatibility
    Given the docx.tracked_changes module
     Then Revision is TrackedChange
      And Insertion, Deletion, and Move are subclasses of Revision


  Scenario: MoveRevision remains importable as an alias of Move
    Given the docx.tracked_changes module
     Then MoveRevision is Move
