Feature: Mutate bookmarks
  In order to revise the bookmarks in a document
  As a developer using python-docx
  I need to delete and rename bookmarks


  Scenario: Bookmark.delete() removes both markers
    Given a document having bookmarks
     When I delete the bookmark named "bm_intro"
     Then len(document.bookmarks) == 2
      And "bm_intro" not in document.bookmarks
      And no bookmarkStart with that id remains in the body
      And no bookmarkEnd with that id remains in the body


  Scenario: Bookmark.delete() works for a cross-paragraph bookmark
    Given a document having bookmarks
     When I delete the bookmark named "bm_span"
     Then len(document.bookmarks) == 2
      And "bm_span" not in document.bookmarks
      And no bookmarkStart with that id remains in the body
      And no bookmarkEnd with that id remains in the body


  Scenario: Deleting one bookmark preserves the others
    Given a document having bookmarks
     When I delete the bookmark named "bm_middle"
     Then len(document.bookmarks) == 2
      And "bm_intro" in document.bookmarks
      And "bm_span" in document.bookmarks


  Scenario: Rename a bookmark
    Given a document having bookmarks
     When I rename the bookmark "bm_intro" to "bm_renamed"
     Then len(document.bookmarks) == 3
      And "bm_renamed" in document.bookmarks
      And "bm_intro" not in document.bookmarks
      And the bookmark "bm_renamed" keeps its original bookmark_id
