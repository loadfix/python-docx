Feature: Read tracked insertions and deletions
  In order to inspect tracked changes a reviewer has made to a document
  As a developer using python-docx
  I need to iterate TrackedChange proxies surfacing authorship and type


  Scenario: Paragraph.tracked_changes lists ins and del children in order
    Given a document with a tracked insertion and deletion in paragraph 1
     Then paragraph.tracked_changes has 2 entries
      And the tracked-change types are ['deletion', 'insertion']
      And the tracked-change authors are ['Bob', 'Alice']
      And the tracked-change texts are ['brown', 'nimble']


  Scenario: TrackedChange exposes author metadata
    Given a document with a tracked insertion and deletion in paragraph 1
     Then the first tracked change author is "Bob"
      And the first tracked change date is a datetime


  Scenario: TrackedChange.type reports "insertion" for w:ins
    Given a document with a tracked insertion and deletion in paragraph 1
     Then tracked_change[1].type == "insertion"


  Scenario: TrackedChange.type reports "deletion" for w:del
    Given a document with a tracked insertion and deletion in paragraph 1
     Then tracked_change[0].type == "deletion"


  Scenario: A paragraph with no tracked changes yields an empty list
    Given the trk-ins-del document
     Then paragraph 3 has no tracked changes


  Scenario: Iterating tracked changes across every paragraph
    Given the trk-ins-del document
     Then iterating every paragraph's tracked_changes yields 3 TrackedChange objects
