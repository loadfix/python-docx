Feature: Add, read, and remove page-watermark shapes on a section
  In order to mark drafts, confidential documents, or add brand imagery
  As a developer using python-docx
  I need Section.add_text_watermark, Section.add_image_watermark, Section.watermark, Section.remove_watermark


  Scenario: Reading a text watermark from a section
    Given a document having a text watermark
     Then section.watermark is a Watermark object
      And section.watermark.type == "text"
      And section.watermark.text == "DRAFT"


  Scenario: Reading an image watermark from a section
    Given a document having an image watermark
     Then section.watermark is a Watermark object
      And section.watermark.type == "image"
      And section.watermark.text is None


  Scenario: Section.watermark is None when no watermark is present
    Given a fresh default document
     Then section.watermark is None


  Scenario: Adding a text watermark replaces any existing one
    Given a fresh default document
     When I add a text watermark with text "DRAFT"
     Then section.watermark.type == "text"
      And section.watermark.text == "DRAFT"


  Scenario: Removing a watermark clears it from the header
    Given a document having a text watermark
     When I call section.remove_watermark()
     Then section.watermark is None


  Scenario: add_text_watermark rejects invalid layout values
    Given a fresh default document
     Then calling add_text_watermark with layout "sideways" raises ValueError
