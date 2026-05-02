Feature: Append the body of one document into another
  In order to merge pre-authored Word content into a target document
  As a developer using python-docx
  I need Document.append_document to import paragraphs, images, and styles


  Scenario: Append a source document into a blank destination
    Given a blank destination document and a source document with content
     When I call dest.append_document(source)
     Then dest has every paragraph from the source
      And dest has the image relationship from the source
      And dest has the Heading 1 style from the source
