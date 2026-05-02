Feature: Mutate endnotes in a document
  In order to modify or remove endnotes already present in a document
  As a developer using python-docx
  I need mutation methods on Endnote objects


  Scenario: Endnote.add_paragraph() without text or style
    Given a document having an endnotes part
     When I assign endnote = the first user endnote in the document
      And I assign paragraph = endnote.add_paragraph()
     Then len(endnote.paragraphs) == 2
      And paragraph.text == ""
      And paragraph.style == "EndnoteText"
      And endnote.paragraphs[-1] == paragraph


  Scenario: Endnote.add_paragraph() specifying text and style
    Given a document having an endnotes part
     When I assign endnote = the first user endnote in the document
      And I assign paragraph = endnote.add_paragraph(text, style)
     Then len(endnote.paragraphs) == 2
      And paragraph.text == text
      And paragraph.style == style
      And endnote.paragraphs[-1] == paragraph


  Scenario: Endnote.clear() empties the endnote but preserves its ref mark
    Given a document having an endnotes part
     When I assign endnote = the first user endnote in the document
      And I assign endnote.add_paragraph("Extra paragraph")
      And I call endnote.clear()
     Then len(endnote.paragraphs) == 1
      And endnote.paragraphs[0].text == ""
      And endnote.paragraphs[0].style.name == "EndnoteText"


  Scenario: Endnote.clear() returns the endnote for fluent chaining
    Given a document having an endnotes part
     When I assign endnote = the first user endnote in the document
     Then endnote.clear() returns endnote


  Scenario: Endnote.delete() removes the endnote from the document
    Given a document having an endnotes part
     When I assign endnote = the first user endnote in the document
      And I call endnote.delete()
     Then len(document.endnotes) == 2
      And iterating document.endnotes yields endnote ids [3, 4]
