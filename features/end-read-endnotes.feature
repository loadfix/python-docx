Feature: Read endnotes from a document
  In order to inspect endnotes present in an existing document
  As a developer using python-docx
  I need read access to the endnotes collection and to each endnote's content


  Scenario: Document.endnotes returns an Endnotes object when a part exists
    Given a document having an endnotes part
     Then document.endnotes is an Endnotes object


  Scenario: Document.endnotes returns an Endnotes object when no part exists
    Given a document having no endnotes part
     Then document.endnotes is an Endnotes object


  Scenario: len(Endnotes) counts only user endnotes
    Given a document having an endnotes part
     Then len(document.endnotes) == 3


  Scenario: len(Endnotes) is zero when no endnotes part exists
    Given a document having no endnotes part
     Then len(document.endnotes) == 0


  Scenario: Endnote.endnote_id is the identifier assigned by Word
    Given a document having an endnotes part
     When I assign endnote = the first user endnote in the document
     Then endnote.endnote_id == 2


  Scenario: Endnote.text gathers the text of the endnote's paragraphs
    Given a document having an endnotes part
     When I assign endnote = the first user endnote in the document
     Then endnote.text == "First endnote citation."


  Scenario: Iteration yields endnotes in document order
    Given a document having an endnotes part
     Then iterating document.endnotes yields endnote ids [2, 3, 4]


  Scenario: Endnote.paragraphs exposes the endnote's paragraph content
    Given a document having an endnotes part
     When I assign endnote = the first user endnote in the document
     Then len(endnote.paragraphs) == 1
      And endnote.paragraphs[0].style.name == "EndnoteText"
