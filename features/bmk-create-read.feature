Feature: Create and read bookmarks
  In order to mark and cross-reference ranges of text in a document
  As a developer using python-docx
  I need to create bookmarks and access the document's bookmarks collection


  Scenario: Access document bookmarks collection
    Given a document having bookmarks
     Then document.bookmarks is a Bookmarks object


  Scenario: Bookmarks.__len__()
    Given a document having bookmarks
     Then len(document.bookmarks) == 3


  Scenario: Bookmarks.__iter__()
    Given a document having bookmarks
     Then iterating document.bookmarks yields 3 Bookmark objects


  Scenario: Bookmarks.get() by name
    Given a document having bookmarks
     When I call document.bookmarks.get("bm_intro")
     Then the result is a Bookmark object named "bm_intro"


  Scenario: Bookmarks.get() returns None for unknown name
    Given a document having bookmarks
     When I call document.bookmarks.get("does_not_exist")
     Then the result is None


  Scenario: Bookmarks containment check by name
    Given a document having bookmarks
     Then "bm_intro" in document.bookmarks
      And "does_not_exist" not in document.bookmarks


  Scenario: Paragraph.add_bookmark() wrapping the whole paragraph
    Given a fresh document with one paragraph of text
     When I assign bookmark = paragraph.add_bookmark("intro")
     Then bookmark.name == "intro"
      And bookmark.bookmark_id == 0
      And len(document.bookmarks) == 1


  Scenario: Paragraph.add_bookmark() wrapping an existing run
    Given a paragraph with three runs
     When I add a bookmark named "middle" around the middle run
     Then bookmark.name == "middle"
      And the bookmark wraps only the middle run
      And len(document.bookmarks) == 1


  Scenario: Paragraph.add_bookmark() across adjacent runs
    Given a paragraph with three runs
     When I add a bookmark named "tail" around the second and third runs
     Then bookmark.name == "tail"
      And the bookmark wraps the last two runs
      And len(document.bookmarks) == 1


  Scenario: Bookmark spanning multiple paragraphs
    Given a document having bookmarks
     When I call document.bookmarks.get("bm_span")
     Then the result is a Bookmark object named "bm_span"
      And the bookmarkStart and bookmarkEnd for "bm_span" are in different paragraphs


  Scenario: Bookmark IDs are unique in the document
    Given a document having bookmarks
     Then every bookmark has a unique bookmark_id
