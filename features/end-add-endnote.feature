Feature: Add an endnote to a document
  In order to add an endnote to a document
  As a developer using python-docx
  I need a way to add an endnote specifying both its content and its reference run


  Scenario: Endnotes.add(run) on a document without an endnotes part
    Given a document having no endnotes part
     When I assign endnote = document.endnotes.add(run)
     Then endnote is an Endnote object
      And endnote.endnote_id == 2
      And len(endnote.paragraphs) == 1
      And endnote.paragraphs[0].style.name == "EndnoteText"
      And len(document.endnotes) == 1


  Scenario: Endnotes.add(run, text) specifying endnote text
    Given a document having an endnotes part
     When I assign endnote = document.endnotes.add(run, "An endnote body")
     Then endnote.text == "An endnote body"
      And len(document.endnotes) == 4


  Scenario: Subsequent Endnotes.add() calls receive successive ids
    Given a document having no endnotes part
     When I add two endnotes to two different runs
     Then the added endnote ids are [2, 3]
      And len(document.endnotes) == 2


  Scenario: Iterating document.endnotes yields only user endnotes
    Given a document having an endnotes part
     Then iterating document.endnotes yields 3 Endnote objects


  Scenario: Iterating a document with no endnotes part yields nothing
    Given a document having no endnotes part
     Then iterating document.endnotes yields 0 Endnote objects
