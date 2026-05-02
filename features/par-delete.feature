Feature: Delete a paragraph
  In order to remove a paragraph from the middle of a document
  As a developer using python-docx
  I need a way to delete an existing paragraph in place


  Scenario: Paragraph.delete() removes the paragraph from the body
    Given a document containing five paragraphs
     When I delete the third paragraph
     Then the document contains four paragraphs
      And the document paragraph text sequence is "Intro, Alpha Beta Gamma, Outro, Tail"


  Scenario: Paragraph.delete() preserves siblings and their formatting
    Given a document containing five paragraphs
     When I delete the third paragraph
     Then the first paragraph has style "Heading 1"
      And the third paragraph has style "Heading 2"


  Scenario: Paragraph.delete() on a detached paragraph is a no-op
    Given a detached paragraph
     When I delete the detached paragraph
     Then no error is raised
